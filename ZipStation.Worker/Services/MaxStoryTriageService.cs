using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using MongoDB.Driver;
using ZipStation.Worker.Entities;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Repositories;

namespace ZipStation.Worker.Services;

/// Pre-creation triage: a single Anthropic call that turns one inbound Discord post into
/// one-or-more proposed kanban cards. Used by DiscordPollingService before card creation.
/// This service does NOT persist cards or tasks — the caller does that with the returned proposals.
public interface IMaxStoryTriageService
{
    Task<TriageResult?> TriagePostAsync(
        Project project,
        DiscordPostContext post,
        CancellationToken cancellationToken = default);
}

public class DiscordPostContext
{
    public string ThreadTitle { get; init; } = string.Empty;
    public string Body { get; init; } = string.Empty;
    public List<string> ForumTags { get; init; } = new();
    public string? AuthorName { get; init; }
    public string SourceName { get; init; } = string.Empty;
    /// Default card type from the source row (built-in name or custom type id). Null = "Auto,
    /// let Max decide" (queued from chunk 1 UX).
    public string? SourceDefaultCardType { get; init; }
}

public class TriageResult
{
    public List<ProposedCard> Cards { get; init; } = new();
    public bool FlaggedQuestion { get; init; }
    public string? QuestionForMaintainer { get; init; }
    public string? QuestionContextExcerpt { get; init; }
    public string? Reasoning { get; init; }
    public string ModelUsed { get; init; } = string.Empty;
    public string RawResponse { get; init; } = string.Empty;
}

public class ProposedCard
{
    public string Title { get; init; } = string.Empty;
    /// A built-in type name (Feature/Bug/Improvement/TechDebt) or, via the source default, a
    /// custom type id. Falls back to the source default if Max gave nothing.
    public string Type { get; init; } = KanbanCardTypes.Bug;
    /// 0=Low, 1=Normal, 2=High, 3=Urgent.
    public int Priority { get; init; } = 1;
    public List<string> Tags { get; init; } = new();
    public string Summary { get; init; } = string.Empty;
    /// Max's cleaned-up rewrite of the original Discord post (basic HTML: p / ul / ol / li / strong / em / code / br).
    /// Empty if Max didn't return one or it failed sanity checks — caller should fall back to the raw post text.
    public string DescriptionHtml { get; init; } = string.Empty;
    /// If set, Max thinks this card duplicates an existing one. Card is still created but a pending
    /// merge_story_duplicate task is attached for human approval — never auto-merge.
    public long? IsDuplicateOfCardNumber { get; init; }
}

public class MaxStoryTriageService : IMaxStoryTriageService
{
    private const string MessagesEndpoint = "https://api.anthropic.com/v1/messages";
    private const string AnthropicVersion = "2023-06-01";
    private const int MaxOutputTokens = 1500;
    private const int MaxRecentStoriesInPrompt = 25;
    private const int MaxInstructionsInPrompt = 20;

    private static readonly HttpClient _httpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(60)
    };

    private readonly KanbanCardRepository _cardRepository;
    private readonly MaxInstructionRepository _instructionRepository;
    private readonly ILogger<MaxStoryTriageService> _logger;
    private readonly MongoDB.Driver.IMongoDatabase _database;
    private readonly AppConfig _appConfig;

    public MaxStoryTriageService(
        KanbanCardRepository cardRepository,
        MaxInstructionRepository instructionRepository,
        MongoDB.Driver.IMongoDatabase database,
        Microsoft.Extensions.Options.IOptions<AppConfig> appConfig,
        ILogger<MaxStoryTriageService> logger)
    {
        _cardRepository = cardRepository;
        _instructionRepository = instructionRepository;
        _database = database;
        _appConfig = appConfig.Value;
        _logger = logger;
    }

    public async Task<TriageResult?> TriagePostAsync(
        Project project,
        DiscordPostContext post,
        CancellationToken cancellationToken = default)
    {
        var maxSettings = project.Settings?.Max;
        if (maxSettings == null || !maxSettings.Enabled || string.IsNullOrEmpty(maxSettings.ApiKeyEncrypted))
        {
            return null;
        }

        var apiKey = EncryptionHelper.Decrypt(maxSettings.ApiKeyEncrypted);
        if (string.IsNullOrEmpty(apiKey))
        {
            _logger.LogWarning("Decrypted Max API key empty for project {ProjectId}; falling back to non-Max card creation", project.Id);
            return null;
        }

        var model = string.IsNullOrWhiteSpace(maxSettings.Model) ? "claude-sonnet-4-6" : maxSettings.Model;

        // Recent stories give Max enough context to spot duplicates.
        var recentStories = await GetRecentNonDoneStoriesAsync(project.Id);
        var instructions = await _instructionRepository.GetByProjectIdAsync(project.Id);

        var systemPrompt = BuildSystemPrompt(maxSettings, instructions, post);
        var userMessage = BuildUserMessage(post, recentStories);

        var rawResponse = await CallAnthropicAsync(apiKey, model, systemPrompt, userMessage, cancellationToken);
        if (rawResponse == null) return null;

        var parsed = ParseTriageResponse(rawResponse);
        if (parsed == null || parsed.Cards == null || parsed.Cards.Count == 0)
        {
            _logger.LogWarning("Max triage returned no cards for thread {ThreadId}; falling back to single-card creation", post.ThreadTitle);
            return null;
        }

        var cards = parsed.Cards
            .Where(c => !string.IsNullOrWhiteSpace(c.Title))
            .Select(c => new ProposedCard
            {
                Title = c.Title!.Trim(),
                Type = NormalizeCardType(c.Type, post.SourceDefaultCardType),
                Priority = NormalizePriority(c.Priority),
                Tags = c.Tags ?? new List<string>(),
                Summary = c.Summary?.Trim() ?? string.Empty,
                DescriptionHtml = SanitizeDescriptionHtml(c.DescriptionHtml),
                IsDuplicateOfCardNumber = c.IsDuplicateOfCardNumber,
            })
            .ToList();

        if (cards.Count == 0) return null;

        return new TriageResult
        {
            Cards = cards,
            FlaggedQuestion = parsed.FlagQuestion,
            QuestionForMaintainer = parsed.QuestionForMaintainer,
            QuestionContextExcerpt = parsed.QuestionContextExcerpt,
            Reasoning = parsed.Reasoning,
            ModelUsed = model,
            RawResponse = rawResponse,
        };
    }

    private async Task<List<KanbanCard>> GetRecentNonDoneStoriesAsync(string projectId)
    {
        // Hand-rolled mongo query rather than adding yet another repo method.
        // Recent non-void cards in the project, capped — used as duplicate-detection context.
        var collection = _database.GetCollection<KanbanCard>(_appConfig.ZipStationMongoDb.Collections.KanbanCards);
        var filter = MongoDB.Driver.Builders<KanbanCard>.Filter.Eq(c => c.ProjectId, projectId)
                   & MongoDB.Driver.Builders<KanbanCard>.Filter.Eq(c => c.IsVoid, false)
                   & MongoDB.Driver.Builders<KanbanCard>.Filter.Eq(c => c.ResolvedOnDateTime, 0);
        return await collection.Find(filter)
            .SortByDescending(c => c.UpdatedOnDateTime)
            .Limit(MaxRecentStoriesInPrompt)
            .ToListAsync();
    }

    private static string BuildSystemPrompt(MaxSettings settings, List<MaxInstruction> instructions, DiscordPostContext post)
    {
        var sb = new StringBuilder();
        sb.AppendLine("You are Max, an engineering triage assistant. Your job is to convert an inbound user post into one or more kanban stories for the maintainer's backlog.");
        sb.AppendLine();
        sb.AppendLine("You do NOT reply to the user. The kanban board is internal. The original post is the source of truth; you cannot ask for clarification.");
        sb.AppendLine();

        if (!string.IsNullOrWhiteSpace(settings.ProjectContext))
        {
            sb.AppendLine("<project_context>");
            sb.AppendLine(settings.ProjectContext);
            sb.AppendLine("</project_context>");
            sb.AppendLine();
        }

        // Pull the enrichment + all-context-applicable instructions; these are general directives the maintainer wrote.
        var applicableInstructions = instructions
            .Where(i => i.Contexts.Contains("enrichment") || i.Contexts.Contains("all"))
            .Take(MaxInstructionsInPrompt)
            .ToList();
        if (applicableInstructions.Count > 0)
        {
            sb.AppendLine("<instructions>");
            foreach (var inst in applicableInstructions)
                sb.AppendLine("- " + inst.Instruction);
            sb.AppendLine("</instructions>");
            sb.AppendLine();
        }

        sb.AppendLine("## Decision rules");
        sb.AppendLine();
        sb.AppendLine("1. **Card count**: Usually one post = one card. Split into multiple cards ONLY when the post clearly describes multiple independent issues (e.g. \"Bug 1: login broken. Bug 2: avatar upload fails.\"). When in doubt, output a single card.");
        sb.AppendLine();
        sb.AppendLine("2. **Type**: classify each card as one of: Feature, Bug, Improvement, TechDebt.");
        if (!string.IsNullOrWhiteSpace(post.SourceDefaultCardType))
            sb.AppendLine($"   - This source's default type is {CardTypeName(post.SourceDefaultCardType)}. Use that unless the post clearly indicates otherwise.");
        else
            sb.AppendLine("   - The source has no default — choose based on the post content alone.");
        sb.AppendLine();
        sb.AppendLine("3. **Priority**: Low / Normal / High / Urgent. Default Normal. Use Urgent only when the post describes a production outage, data loss, or security issue.");
        sb.AppendLine();
        sb.AppendLine("4. **Duplicates**: If the post obviously duplicates an existing non-done story listed in <available_stories>, set is_duplicate_of_card_number to that card's number on the proposed card. Do NOT skip the card — produce it anyway, the maintainer will approve the merge. Only flag clear duplicates, not loosely-related stories.");
        sb.AppendLine();
        sb.AppendLine("5. **Title**: 5-12 words, imperative or noun phrase. Strip greetings, apologies, and conversational filler (\"hey guys\", \"please help\", \"sorry to bother\"). Make it scannable on the board.");
        sb.AppendLine();
        sb.AppendLine("6. **Tags**: optional. 0-3 short lowercase tags (e.g. \"mobile\", \"login\", \"perf\"). Only obvious tags.");
        sb.AppendLine();
        sb.AppendLine("7. **Summary**: 1 sentence, max ~140 chars. The maintainer reads this at a glance.");
        sb.AppendLine();
        sb.AppendLine("8. **Description (description_html)**: rewrite the original post into a clean, scannable kanban description that's easy to read at a glance. Specifically:");
        sb.AppendLine("   - Strip preamble, hedging, apologies, and conversational filler. The user said \"I'm not sure if this is possible but I'd really appreciate it if...\" — you write \"Add X.\"");
        sb.AppendLine("   - Lead with what the user wants (the ask), then any context they provided.");
        sb.AppendLine("   - Preserve all concrete facts verbatim: specific names, error messages, version numbers, URLs, reproduction steps, code snippets. Don't paraphrase identifiers.");
        sb.AppendLine("   - Use short paragraphs. Use a `<ul>` of `<li>` bullets when the user listed several distinct points; don't fabricate bullets where the original was one continuous thought.");
        sb.AppendLine("   - Allowed tags ONLY: `<p>`, `<ul>`, `<ol>`, `<li>`, `<strong>`, `<em>`, `<code>`, `<br>`. No links, images, scripts, styles, or attributes. The original post is reachable via the story's external-source link, so don't add \"original post:\" callouts.");
        sb.AppendLine("   - Don't invent information the user didn't provide. Don't speculate about implementation.");
        sb.AppendLine("   - If the original post is already short and clean, just lightly format it — don't pad.");
        sb.AppendLine();
        sb.AppendLine("9. **Flag a question** only when the post is unintelligible or asks for support rather than reporting an issue. Do not flag for every ambiguity.");
        sb.AppendLine();

        sb.AppendLine("## Output");
        sb.AppendLine();
        sb.AppendLine("Output a single JSON object, no markdown fences, no prose. Schema:");
        sb.AppendLine("{");
        sb.AppendLine("  \"cards\": [");
        sb.AppendLine("    {");
        sb.AppendLine("      \"title\": \"string\",");
        sb.AppendLine("      \"type\": \"Feature\" | \"Bug\" | \"Improvement\" | \"TechDebt\",");
        sb.AppendLine("      \"priority\": \"Low\" | \"Normal\" | \"High\" | \"Urgent\",");
        sb.AppendLine("      \"tags\": [\"string\"],");
        sb.AppendLine("      \"summary\": \"string\",");
        sb.AppendLine("      \"description_html\": \"string\",");
        sb.AppendLine("      \"is_duplicate_of_card_number\": number | null");
        sb.AppendLine("    }");
        sb.AppendLine("  ],");
        sb.AppendLine("  \"flag_question\": boolean,");
        sb.AppendLine("  \"question_for_maintainer\": \"string\" | null,");
        sb.AppendLine("  \"question_context_excerpt\": \"string\" | null,");
        sb.AppendLine("  \"reasoning\": \"string\"");
        sb.AppendLine("}");

        return sb.ToString();
    }

    private static string BuildUserMessage(DiscordPostContext post, List<KanbanCard> recentStories)
    {
        var sb = new StringBuilder();

        sb.AppendLine("<available_stories>");
        if (recentStories.Count == 0)
        {
            sb.AppendLine("(none)");
        }
        else
        {
            foreach (var s in recentStories)
            {
                sb.AppendLine($"- STR-{s.CardNumber} [{CardTypeName(s.Type)}]: {s.Title}");
            }
        }
        sb.AppendLine("</available_stories>");
        sb.AppendLine();

        sb.AppendLine("<discord_post>");
        sb.AppendLine($"Source: {post.SourceName}");
        if (!string.IsNullOrWhiteSpace(post.AuthorName))
            sb.AppendLine($"Author: {post.AuthorName}");
        if (post.ForumTags.Count > 0)
            sb.AppendLine($"Forum tags: {string.Join(", ", post.ForumTags)}");
        sb.AppendLine($"Title: {post.ThreadTitle}");
        sb.AppendLine("Body:");
        sb.AppendLine(string.IsNullOrWhiteSpace(post.Body) ? "(empty body)" : post.Body);
        sb.AppendLine("</discord_post>");

        return sb.ToString();
    }

    private async Task<string?> CallAnthropicAsync(string apiKey, string model, string systemPrompt, string userMessage, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, MessagesEndpoint);
        request.Headers.TryAddWithoutValidation("x-api-key", apiKey);
        request.Headers.TryAddWithoutValidation("anthropic-version", AnthropicVersion);
        request.Content = JsonContent.Create(new
        {
            model,
            max_tokens = MaxOutputTokens,
            temperature = 0,
            system = systemPrompt,
            messages = new[] { new { role = "user", content = userMessage } }
        });

        using var response = await _httpClient.SendAsync(request, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            _logger.LogWarning("Anthropic returned {Status} for story triage: {Body}", (int)response.StatusCode, body);
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("content", out var content) || content.GetArrayLength() == 0)
                return null;
            var first = content[0];
            if (!first.TryGetProperty("text", out var text)) return null;
            return text.GetString();
        }
        catch (JsonException ex)
        {
            _logger.LogWarning(ex, "Failed to parse Anthropic envelope for story triage");
            return null;
        }
    }

    private static ParsedTriage? ParseTriageResponse(string json)
    {
        try
        {
            var trimmed = json.Trim();
            if (trimmed.StartsWith("```"))
            {
                var firstNewline = trimmed.IndexOf('\n');
                if (firstNewline > 0) trimmed = trimmed.Substring(firstNewline + 1);
                if (trimmed.EndsWith("```")) trimmed = trimmed.Substring(0, trimmed.Length - 3);
                trimmed = trimmed.Trim();
            }

            return JsonSerializer.Deserialize<ParsedTriage>(trimmed, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
                NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.AllowReadingFromString,
            });
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static string NormalizeCardType(string? raw, string? sourceDefault)
    {
        if (string.IsNullOrWhiteSpace(raw)) return sourceDefault ?? KanbanCardTypes.Bug;
        return raw.Trim().ToLowerInvariant() switch
        {
            "feature" => KanbanCardTypes.Feature,
            "bug" => KanbanCardTypes.Bug,
            "improvement" => KanbanCardTypes.Improvement,
            "techdebt" or "tech_debt" or "tech-debt" => KanbanCardTypes.TechDebt,
            _ => sourceDefault ?? KanbanCardTypes.Bug,
        };
    }

    private static int NormalizePriority(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return 1;
        return raw.Trim().ToLowerInvariant() switch
        {
            "low" => 0,
            "normal" => 1,
            "high" => 2,
            "urgent" => 3,
            _ => 1,
        };
    }

    /// Story type is already a string (built-in name or custom id). Built-ins are shown as-is;
    /// a custom id is shown verbatim (rare — only when a source default points at a custom type).
    private static string CardTypeName(string? type) =>
        string.IsNullOrWhiteSpace(type) ? "Unknown" : type;

    private class ParsedTriage
    {
        public List<ParsedTriageCard>? Cards { get; set; }
        public bool FlagQuestion { get; set; }
        public string? QuestionForMaintainer { get; set; }
        public string? QuestionContextExcerpt { get; set; }
        public string? Reasoning { get; set; }
    }

    private class ParsedTriageCard
    {
        public string? Title { get; set; }
        public string? Type { get; set; }
        public string? Priority { get; set; }
        public List<string>? Tags { get; set; }
        public string? Summary { get; set; }
        public string? DescriptionHtml { get; set; }
        public long? IsDuplicateOfCardNumber { get; set; }
    }

    /// Belt-and-suspenders: the prompt restricts Max to a safe HTML subset, but strip anything
    /// that could be dangerous in case Max ignores instructions or a hostile post tries to
    /// inject via prompt-content. Returns empty when nothing usable remains so the caller can
    /// fall back to the raw post text.
    private static readonly System.Text.RegularExpressions.Regex _scriptyTagPattern = new(
        @"<\s*/?\s*(script|style|iframe|object|embed|svg|math|link|meta|form|input|button|textarea|select|video|audio|source|track)\b[^>]*>",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
    private static readonly System.Text.RegularExpressions.Regex _eventHandlerAttrPattern = new(
        @"\s+on[a-z]+\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)",
        System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string SanitizeDescriptionHtml(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        var trimmed = raw.Trim();
        trimmed = _scriptyTagPattern.Replace(trimmed, "");
        trimmed = _eventHandlerAttrPattern.Replace(trimmed, "");
        return trimmed;
    }
}
