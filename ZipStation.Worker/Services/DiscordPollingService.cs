using System.Globalization;
using System.Net.Http.Headers;
using System.Text.Json;
using MongoDB.Bson;
using ZipStation.Worker.Entities;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Repositories;

namespace ZipStation.Worker.Services;

public interface IDiscordPollingService
{
    Task PollAllProjectsAsync(CancellationToken ct);
    Task PollProjectAsync(string projectId, CancellationToken ct);
}

/// Forum-channel-first Discord intake. For each enabled source:
///   1. List active threads in the guild (well-documented endpoint).
///   2. Filter to threads whose parent_id matches the source channel.
///   3. Keep threads with snowflake id &gt; source.LastSeenId.
///   4. For each new thread, fetch the starter message (whose id == thread id),
///      build a KanbanCard, and persist it. Discord URL is pre-built so the SPA
///      just renders a link.
public class DiscordPollingService : IDiscordPollingService
{
    private const string DiscordApiBase = "https://discord.com/api/v10";
    private const int DiscordSourceType = 0; // see KanbanCardExternalSource.Type comment
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(30) };
    // Discord's API uses snake_case (parent_id, applied_tags, owner_id, available_tags).
    // Case-insensitive alone doesn't bridge that to PascalCase properties — every snake-cased
    // field would silently deserialize to null. SnakeCaseLower naming policy is the fix.
    private static readonly JsonSerializerOptions _json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        PropertyNameCaseInsensitive = true,
    };

    private readonly ILogger<DiscordPollingService> _logger;
    private readonly ProjectRepository _projectRepository;
    private readonly KanbanBoardRepository _boardRepository;
    private readonly KanbanCardRepository _cardRepository;
    private readonly KanbanCardNumberCounterRepository _cardNumberCounter;
    private readonly IMaxStoryTriageService _maxTriage;
    private readonly MaxStoryEnrichmentRepository _storyEnrichmentRepository;
    private readonly MaxTaskRepository _maxTaskRepository;
    private readonly MaxQuestionRepository _maxQuestionRepository;

    static DiscordPollingService()
    {
        // Discord requires a descriptive User-Agent; default .NET UA can be rate-limited.
        _http.DefaultRequestHeaders.UserAgent.ParseAdd("ZipStationWorker (https://zipstation.dev, 1.0)");
    }

    public DiscordPollingService(
        ILogger<DiscordPollingService> logger,
        ProjectRepository projectRepository,
        KanbanBoardRepository boardRepository,
        KanbanCardRepository cardRepository,
        KanbanCardNumberCounterRepository cardNumberCounter,
        IMaxStoryTriageService maxTriage,
        MaxStoryEnrichmentRepository storyEnrichmentRepository,
        MaxTaskRepository maxTaskRepository,
        MaxQuestionRepository maxQuestionRepository)
    {
        _logger = logger;
        _projectRepository = projectRepository;
        _boardRepository = boardRepository;
        _cardRepository = cardRepository;
        _cardNumberCounter = cardNumberCounter;
        _maxTriage = maxTriage;
        _storyEnrichmentRepository = storyEnrichmentRepository;
        _maxTaskRepository = maxTaskRepository;
        _maxQuestionRepository = maxQuestionRepository;
    }

    public async Task PollAllProjectsAsync(CancellationToken ct)
    {
        var projects = await _projectRepository.GetAllWithDiscordEnabledAsync();
        foreach (var project in projects)
        {
            if (ct.IsCancellationRequested) break;
            try
            {
                await PollProjectInternalAsync(project, ct);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Discord poll failed for project {ProjectId}", project.Id);
            }
        }
    }

    public async Task PollProjectAsync(string projectId, CancellationToken ct)
    {
        var project = await _projectRepository.GetByIdAsync(projectId);
        if (project?.Settings?.Discord is not { Enabled: true }) return;
        await PollProjectInternalAsync(project, ct);
    }

    private async Task PollProjectInternalAsync(Project project, CancellationToken ct)
    {
        var discord = project.Settings?.Discord;
        if (discord is null || !discord.Enabled) return;

        var token = EncryptionHelper.Decrypt(discord.BotTokenEncrypted ?? "");
        if (string.IsNullOrWhiteSpace(token))
        {
            _logger.LogWarning("Discord enabled for project {ProjectId} but bot token is empty after decrypt", project.Id);
            return;
        }

        var board = await _boardRepository.GetByProjectIdAsync(project.Id);
        if (board is null || board.Columns.Count == 0)
        {
            _logger.LogWarning("Discord poll skipped for project {ProjectId}: no kanban board / columns", project.Id);
            return;
        }
        // Automated intake lands in the board's configured intake column (falls back to the
        // lowest-position column when unset). Avoids silently dropping cards into whatever
        // column happens to sit at position 0 — e.g. an "Obsolete" column added in front.
        var intakeColumnId = board.ResolveIntakeColumnId();

        foreach (var source in discord.Sources.Where(s => s.Enabled))
        {
            if (ct.IsCancellationRequested) break;
            if (string.IsNullOrWhiteSpace(source.GuildId) || string.IsNullOrWhiteSpace(source.ChannelId)) continue;

            try
            {
                var newThreads = await FetchNewForumThreadsAsync(token, source, ct);
                if (newThreads.Count == 0) continue;

                // Resolve forum applied_tags IDs → names once per poll cycle for this source.
                // Channel tags rarely change; one fetch covers every thread we process below.
                var forumTagNames = source.IsForum
                    ? await FetchForumTagNamesAsync(token, source.ChannelId, ct)
                    : new Dictionary<string, string>();

                string maxIdSeen = source.LastSeenId ?? "0";
                int created = 0;
                int skippedExisting = 0;
                foreach (var thread in newThreads.OrderBy(t => t.Id, StringComparer.Ordinal))
                {
                    if (ct.IsCancellationRequested) break;

                    var existing = await _cardRepository.GetByExternalMessageIdAsync(project.Id, DiscordSourceType, thread.Id);
                    if (existing is not null)
                    {
                        skippedExisting++;
                        if (CompareSnowflake(thread.Id, maxIdSeen) > 0) maxIdSeen = thread.Id;
                        continue;
                    }

                    var starter = await TryFetchStarterMessageAsync(token, thread.Id, ct);
                    await ProcessNewThreadAsync(project, board.Id, intakeColumnId, source, thread, starter, forumTagNames, ct);
                    created++;

                    if (CompareSnowflake(thread.Id, maxIdSeen) > 0) maxIdSeen = thread.Id;
                }

                _logger.LogInformation(
                    "Discord poll for source {SourceName}: processed {New} thread(s) — {Created} new card(s), {Skipped} already-imported skipped",
                    source.Name, newThreads.Count, created, skippedExisting);

                if (!string.Equals(maxIdSeen, source.LastSeenId, StringComparison.Ordinal))
                {
                    await _projectRepository.UpdateDiscordSourceCursorAsync(project.Id, source.Id, maxIdSeen);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Discord source {SourceName} (channel {ChannelId}) failed in project {ProjectId}",
                    source.Name, source.ChannelId, project.Id);
            }
        }
    }

    private async Task<List<DiscordThread>> FetchNewForumThreadsAsync(string token, DiscordSource source, CancellationToken ct)
    {
        // GET /guilds/{guild.id}/threads/active returns all active threads in the guild.
        // We filter to those whose parent_id matches the source channel.
        var request = new HttpRequestMessage(HttpMethod.Get, $"{DiscordApiBase}/guilds/{source.GuildId}/threads/active");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bot", token);

        using var response = await _http.SendAsync(request, ct);
        if (!response.IsSuccessStatusCode)
        {
            var body = await response.Content.ReadAsStringAsync(ct);
            throw new InvalidOperationException(
                $"Discord active-threads fetch returned {(int)response.StatusCode}: {body.Substring(0, Math.Min(body.Length, 500))}");
        }

        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        var payload = await JsonSerializer.DeserializeAsync<ActiveThreadsResponse>(stream, _json, ct);
        var all = payload?.Threads ?? new List<DiscordThread>();

        var matchedChannel = all.Where(t => t.ParentId == source.ChannelId).ToList();
        var cursor = source.LastSeenId;
        var newThreads = matchedChannel
            .Where(t => string.IsNullOrEmpty(cursor) || CompareSnowflake(t.Id, cursor) > 0)
            .ToList();

        // Visibility for the silent-no-match path. Without this, "Sync now" looks broken when
        // it's actually working — wrong channel id, no posts yet, archived threads, etc. all
        // produce zero work and zero logs.
        _logger.LogInformation(
            "Discord fetch for source {SourceName}: {TotalActive} active thread(s) in guild, {MatchedChannel} in channel {ChannelId}, {NewToProcess} new since cursor {Cursor}",
            source.Name, all.Count, matchedChannel.Count, source.ChannelId, newThreads.Count, source.LastSeenId ?? "(none)");

        // When zero threads match the configured channel but the guild does have active threads,
        // dump the unique parent_ids so the maintainer can see which channels they actually live in
        // and confirm whether they pasted the wrong id.
        if (matchedChannel.Count == 0 && all.Count > 0)
        {
            var parents = all
                .Where(t => !string.IsNullOrEmpty(t.ParentId))
                .GroupBy(t => t.ParentId!)
                .Select(g => $"{g.Key}({g.Count()})")
                .Take(10)
                .ToList();
            _logger.LogInformation(
                "Discord fetch for source {SourceName}: active threads live in parent channels: {ParentIds}",
                source.Name, string.Join(", ", parents));
        }

        return newThreads;
    }

    /// Forum channels expose tag definitions on the channel object itself. We fetch them once
    /// per poll cycle so we can swap snowflake tag IDs on each thread for the human-readable
    /// labels Max and the SPA actually want to display.
    private async Task<Dictionary<string, string>> FetchForumTagNamesAsync(string token, string channelId, CancellationToken ct)
    {
        try
        {
            var request = new HttpRequestMessage(HttpMethod.Get, $"{DiscordApiBase}/channels/{channelId}");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bot", token);

            using var response = await _http.SendAsync(request, ct);
            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Discord channel fetch returned {Status} for channel {ChannelId}; forum tags will not resolve",
                    (int)response.StatusCode, channelId);
                return new Dictionary<string, string>();
            }

            await using var stream = await response.Content.ReadAsStreamAsync(ct);
            var channel = await JsonSerializer.DeserializeAsync<DiscordChannel>(stream, _json, ct);
            var tags = channel?.AvailableTags ?? new List<DiscordForumTag>();
            return tags
                .Where(t => !string.IsNullOrEmpty(t.Id) && !string.IsNullOrEmpty(t.Name))
                .ToDictionary(t => t.Id, t => t.Name!);
        }
        catch (Exception ex)
        {
            // Tag resolution is a nice-to-have. If Discord returns garbage we just skip the labels.
            _logger.LogWarning(ex, "Failed to fetch forum tag definitions for channel {ChannelId}", channelId);
            return new Dictionary<string, string>();
        }
    }

    private async Task<DiscordMessage?> TryFetchStarterMessageAsync(string token, string threadId, CancellationToken ct)
    {
        // In a forum thread, the starter message has the same id as the thread itself.
        var request = new HttpRequestMessage(HttpMethod.Get, $"{DiscordApiBase}/channels/{threadId}/messages/{threadId}");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bot", token);

        using var response = await _http.SendAsync(request, ct);
        if (!response.IsSuccessStatusCode)
        {
            _logger.LogWarning("Discord starter-message fetch returned {Status} for thread {ThreadId}", (int)response.StatusCode, threadId);
            return null;
        }

        await using var stream = await response.Content.ReadAsStreamAsync(ct);
        return await JsonSerializer.DeserializeAsync<DiscordMessage>(stream, _json, ct);
    }

    /// Convert one new Discord thread into one-or-more kanban cards, optionally enriched by Max.
    private async Task ProcessNewThreadAsync(
        Project project, string boardId, string firstColumnId,
        DiscordSource source, DiscordThread thread, DiscordMessage? starter,
        IReadOnlyDictionary<string, string> forumTagNames,
        CancellationToken ct)
    {
        var bodyText = starter?.Content ?? string.Empty;
        var descriptionHtml = string.IsNullOrWhiteSpace(bodyText)
            ? "<p><em>(no body)</em></p>"
            : "<p>" + System.Net.WebUtility.HtmlEncode(bodyText).Replace("\n", "<br/>") + "</p>";
        var fallbackTitle = !string.IsNullOrWhiteSpace(thread.Name) ? thread.Name : Truncate(bodyText, 80);
        if (string.IsNullOrWhiteSpace(fallbackTitle)) fallbackTitle = "Discord post";

        var resolvedTags = (thread.AppliedTags ?? new List<string>())
            .Select(id => forumTagNames.TryGetValue(id, out var name) ? name : id)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .ToList();

        // Triage call (Max) — null when Max disabled / no key / API failure. Falls back to single-card path.
        var post = new DiscordPostContext
        {
            ThreadTitle = thread.Name ?? string.Empty,
            Body = bodyText,
            ForumTags = resolvedTags,
            AuthorName = starter?.Author?.Username,
            SourceName = string.IsNullOrEmpty(source.Name) ? $"#{source.ChannelId}" : source.Name,
            SourceDefaultCardType = source.DefaultCardType,
        };
        var triage = await _maxTriage.TriagePostAsync(project, post, ct);

        if (triage is null || triage.Cards.Count == 0)
        {
            // Fallback path: Max is disabled, failed, or returned nothing.
            // If the source has no default type ("Auto"), there's no Max to make the call — fall back to Bug.
            var card = await CreateCardCoreAsync(
                project, boardId, firstColumnId, source, thread, starter,
                title: fallbackTitle,
                type: source.DefaultCardType ?? KanbanCardTypes.Bug,
                priority: 1,
                tags: new(),
                descriptionHtml: descriptionHtml,
                forumTagNames: resolvedTags);
            _logger.LogInformation("Discord intake (no Max): created STR-{CardNumber} from thread {ThreadId}", card.CardNumber, thread.Id);
            return;
        }

        _logger.LogInformation(
            "Discord intake (Max): thread {ThreadId} → {CardCount} card(s){QuestionFlag}",
            thread.Id, triage.Cards.Count, triage.FlaggedQuestion ? " (question flagged)" : "");

        string? firstCardId = null;
        foreach (var proposed in triage.Cards)
        {
            // Prefer Max's cleaned-up description over the raw post body. Raw body remains
            // accessible via the card's ExternalSources deep-link, so nothing is lost.
            var cardDescription = !string.IsNullOrWhiteSpace(proposed.DescriptionHtml)
                ? proposed.DescriptionHtml
                : descriptionHtml;

            var card = await CreateCardCoreAsync(
                project, boardId, firstColumnId, source, thread, starter,
                title: proposed.Title,
                type: proposed.Type,
                priority: proposed.Priority,
                tags: proposed.Tags,
                descriptionHtml: cardDescription,
                forumTagNames: resolvedTags);
            firstCardId ??= card.Id;

            await PersistStoryEnrichmentAsync(card, proposed, triage);

            if (proposed.IsDuplicateOfCardNumber.HasValue)
            {
                await CreateDuplicateTaskAsync(project, card, proposed.IsDuplicateOfCardNumber.Value);
            }
        }

        if (triage.FlaggedQuestion && !string.IsNullOrWhiteSpace(triage.QuestionForMaintainer) && firstCardId is not null)
        {
            await _maxQuestionRepository.CreateAsync(new MaxQuestion
            {
                CompanyId = project.CompanyId,
                ProjectId = project.Id,
                SourceStoryId = firstCardId,
                Question = triage.QuestionForMaintainer!,
                ContextExcerpt = triage.QuestionContextExcerpt,
                Status = "pending",
            });
        }
    }

    private async Task<KanbanCard> CreateCardCoreAsync(
        Project project, string boardId, string firstColumnId,
        DiscordSource source, DiscordThread thread, DiscordMessage? starter,
        string title, string type, int priority, List<string> tags, string descriptionHtml,
        List<string> forumTagNames)
    {
        var cardNumber = await _cardNumberCounter.GetNextCardNumberAsync(project.Id);
        var maxPosition = await _cardRepository.GetMaxPositionInColumnAsync(boardId, firstColumnId);

        var card = new KanbanCard
        {
            Id = ObjectId.GenerateNewId().ToString(),
            CompanyId = project.CompanyId,
            ProjectId = project.Id,
            BoardId = boardId,
            CardNumber = cardNumber,
            ColumnId = firstColumnId,
            Position = maxPosition + 1,
            Title = Truncate(title, 200),
            DescriptionHtml = descriptionHtml,
            Type = type,
            Priority = priority,
            Tags = tags,
            LinkedTicketIds = new(),
            ExternalSources = new List<KanbanCardExternalSource>
            {
                new()
                {
                    Type = DiscordSourceType,
                    Url = $"https://discord.com/channels/{source.GuildId}/{thread.Id}",
                    GuildId = source.GuildId,
                    ChannelId = source.ChannelId,
                    ThreadId = thread.Id,
                    MessageId = thread.Id,
                    ThreadTitle = thread.Name,
                    ForumTags = forumTagNames,
                    AuthorName = starter?.Author?.Username,
                    AuthorExternalId = starter?.Author?.Id ?? thread.OwnerId,
                }
            }
        };

        return await _cardRepository.CreateAsync(card);
    }

    /// Seed enrichment so the SPA inline panel has something to show as soon as the card opens —
    /// avoids a "no enrichment yet, click Sparkles" empty state on every Max-created card.
    private async Task PersistStoryEnrichmentAsync(KanbanCard card, ProposedCard proposed, TriageResult triage)
    {
        var category = CategoryForCardType(card.Type);
        var suggestedAction = proposed.IsDuplicateOfCardNumber.HasValue ? "merge_story_duplicate" : "no_action";

        await _storyEnrichmentRepository.UpsertAsync(new MaxStoryEnrichment
        {
            CompanyId = card.CompanyId,
            ProjectId = card.ProjectId,
            StoryId = card.Id,
            Status = "complete",
            Category = category,
            Summary = proposed.Summary,
            Confidence = 0.8, // Triage prompt is high-precision; no separate confidence in triage output yet.
            Tags = proposed.Tags,
            SuggestedActionType = suggestedAction,
            Reasoning = triage.Reasoning,
            FlaggedQuestion = triage.FlaggedQuestion,
            Model = triage.ModelUsed,
            RawResponse = triage.RawResponse,
        });
    }

    private async Task CreateDuplicateTaskAsync(Project project, KanbanCard sourceCard, long duplicateOfCardNumber)
    {
        // Resolve the proposed duplicate target. If it's not findable, log and skip — no point creating
        // a merge task that points nowhere.
        var target = await _cardRepository.GetByCardNumberAsync(project.Id, duplicateOfCardNumber);
        if (target is null)
        {
            _logger.LogWarning(
                "Max proposed merge of STR-{Source} into STR-{Target} but target not found in project {ProjectId}",
                sourceCard.CardNumber, duplicateOfCardNumber, project.Id);
            return;
        }

        await _maxTaskRepository.CreateAsync(new MaxTask
        {
            CompanyId = project.CompanyId,
            ProjectId = project.Id,
            TicketId = string.Empty,
            StoryId = sourceCard.Id,
            Type = "merge_story_duplicate",
            Status = "pending",
            Confidence = 0.8,
            Details = new MaxTaskDetails
            {
                DuplicateOfStoryId = target.Id,
                DuplicateOfStoryCardNumber = target.CardNumber,
                DuplicateOfStoryTitle = target.Title,
            }
        });
    }

    private static string CategoryForCardType(string? type) => type switch
    {
        KanbanCardTypes.Feature => "feature",
        KanbanCardTypes.Bug => "bug",
        KanbanCardTypes.Improvement => "improvement",
        KanbanCardTypes.TechDebt => "tech_debt",
        _ => "unclear",
    };

    private static int CompareSnowflake(string a, string b)
    {
        // Snowflakes fit in ulong; lexical compare works only if same length, so parse.
        if (!ulong.TryParse(a, NumberStyles.None, CultureInfo.InvariantCulture, out var av)) return string.CompareOrdinal(a, b);
        if (!ulong.TryParse(b, NumberStyles.None, CultureInfo.InvariantCulture, out var bv)) return string.CompareOrdinal(a, b);
        return av.CompareTo(bv);
    }

    private static string Truncate(string s, int max)
    {
        if (string.IsNullOrEmpty(s)) return s;
        return s.Length <= max ? s : s.Substring(0, max);
    }

    // ----- Discord DTOs (only the fields we need) -----

    private class ActiveThreadsResponse
    {
        public List<DiscordThread> Threads { get; set; } = new();
    }

    private class DiscordThread
    {
        public string Id { get; set; } = string.Empty;
        public string? Name { get; set; }
        public string? ParentId { get; set; }
        public string? OwnerId { get; set; }
        public List<string>? AppliedTags { get; set; }
    }

    private class DiscordMessage
    {
        public string Id { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public DiscordUser? Author { get; set; }
    }

    private class DiscordUser
    {
        public string Id { get; set; } = string.Empty;
        public string Username { get; set; } = string.Empty;
    }

    private class DiscordChannel
    {
        public string Id { get; set; } = string.Empty;
        public List<DiscordForumTag>? AvailableTags { get; set; }
    }

    private class DiscordForumTag
    {
        public string Id { get; set; } = string.Empty;
        public string? Name { get; set; }
    }
}
