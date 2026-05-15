using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using ZipStation.Worker.Entities;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Repositories;

namespace ZipStation.Worker.Services;

public interface IMaxEnrichmentService
{
    Task<MaxTicketEnrichment?> EnrichTicketAsync(string ticketId, CancellationToken cancellationToken = default);
}

public class MaxEnrichmentService : IMaxEnrichmentService
{
    private const string MessagesEndpoint = "https://api.anthropic.com/v1/messages";
    private const string AnthropicVersion = "2023-06-01";
    private const int MaxExampleRepliesInPrompt = 10;
    private const int MaxOpenTicketsInPrompt = 30;
    private const int MaxOutputTokens = 2000;
    private const int CustomerSourceValue = 0; // MessageSource.Customer

    private static readonly HttpClient _httpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(60)
    };

    private readonly ProjectRepository _projectRepository;
    private readonly TicketRepository _ticketRepository;
    private readonly TicketMessageRepository _ticketMessageRepository;
    private readonly MaxInstructionRepository _instructionRepository;
    private readonly MaxExampleReplyRepository _exampleReplyRepository;
    private readonly MaxTicketEnrichmentRepository _enrichmentRepository;
    private readonly MaxTaskRepository _taskRepository;
    private readonly MaxQuestionRepository _questionRepository;
    private readonly ILogger<MaxEnrichmentService> _logger;

    public MaxEnrichmentService(
        ProjectRepository projectRepository,
        TicketRepository ticketRepository,
        TicketMessageRepository ticketMessageRepository,
        MaxInstructionRepository instructionRepository,
        MaxExampleReplyRepository exampleReplyRepository,
        MaxTicketEnrichmentRepository enrichmentRepository,
        MaxTaskRepository taskRepository,
        MaxQuestionRepository questionRepository,
        ILogger<MaxEnrichmentService> logger)
    {
        _projectRepository = projectRepository;
        _ticketRepository = ticketRepository;
        _ticketMessageRepository = ticketMessageRepository;
        _instructionRepository = instructionRepository;
        _exampleReplyRepository = exampleReplyRepository;
        _enrichmentRepository = enrichmentRepository;
        _taskRepository = taskRepository;
        _questionRepository = questionRepository;
        _logger = logger;
    }

    public async Task<MaxTicketEnrichment?> EnrichTicketAsync(string ticketId, CancellationToken cancellationToken = default)
    {
        Ticket? ticket = null;
        try
        {
            ticket = await _ticketRepository.GetByIdAsync(ticketId);
            if (ticket == null)
            {
                _logger.LogWarning("EnrichTicket called for missing ticket {TicketId}", ticketId);
                return null;
            }

            var project = await _projectRepository.GetByIdAsync(ticket.ProjectId);
            if (project?.Settings?.Max == null || !project.Settings.Max.Enabled || string.IsNullOrEmpty(project.Settings.Max.ApiKeyEncrypted))
            {
                return null;
            }

            var apiKey = EncryptionHelper.Decrypt(project.Settings.Max.ApiKeyEncrypted);
            if (string.IsNullOrEmpty(apiKey))
            {
                _logger.LogWarning("Decrypted Max API key empty for project {ProjectId}", ticket.ProjectId);
                return null;
            }

            await SetEnrichmentStatusAsync(ticket, "processing");

            var model = string.IsNullOrWhiteSpace(project.Settings.Max.Model) ? "claude-sonnet-4-6" : project.Settings.Max.Model;
            var instructions = await _instructionRepository.GetByProjectIdAsync(ticket.ProjectId);
            var examples = await _exampleReplyRepository.GetByProjectIdAsync(ticket.ProjectId);
            var messages = await _ticketMessageRepository.GetByTicketIdAsync(ticketId);
            var latestCustomerMessage = messages.LastOrDefault(m => m.Source == CustomerSourceValue);
            var openTickets = await _ticketRepository.GetRecentOpenByProjectIdAsync(ticket.ProjectId, MaxOpenTicketsInPrompt, ticketId);
            var openEnrichments = await _enrichmentRepository.GetRecentByProjectIdAsync(ticket.ProjectId, MaxOpenTicketsInPrompt);
            var enrichmentByTicketId = openEnrichments.ToDictionary(e => e.TicketId);

            var systemPrompt = BuildSystemPrompt(project.Settings.Max, instructions, examples);
            var userMessage = BuildUserMessage(ticket, latestCustomerMessage, openTickets, enrichmentByTicketId);

            var rawResponse = await CallAnthropicAsync(apiKey, model, systemPrompt, userMessage, cancellationToken);
            if (rawResponse == null)
            {
                await SetEnrichmentStatusAsync(ticket, "failed");
                return null;
            }

            var parsed = ParseEnrichmentResponse(rawResponse);
            if (parsed == null)
            {
                _logger.LogWarning("Failed to parse Max enrichment response for ticket {TicketId}", ticketId);
                await SetEnrichmentStatusAsync(ticket, "failed");
                return null;
            }

            if (parsed.SuggestedAction?.Draft != null && MentionsEmDashes(project.Settings.Max.ToneAvoid))
            {
                parsed.SuggestedAction.Draft = StripEmDashes(parsed.SuggestedAction.Draft);
            }

            var existing = await _enrichmentRepository.GetByTicketIdAsync(ticketId);
            var enrichment = existing ?? new MaxTicketEnrichment
            {
                CompanyId = ticket.CompanyId,
                ProjectId = ticket.ProjectId,
                TicketId = ticketId,
            };

            enrichment.Status = "complete";
            enrichment.Category = parsed.Category ?? "unsure";
            enrichment.Summary = parsed.Summary ?? string.Empty;
            enrichment.Confidence = parsed.Confidence;
            enrichment.DuplicateOfTicketId = parsed.DuplicateOf;
            enrichment.RelatedTicketIds = parsed.RelatedIds ?? new();
            enrichment.Platform = parsed.Platform ?? "unknown";
            enrichment.Tags = parsed.Tags ?? new();
            enrichment.SuggestedActionType = parsed.SuggestedAction?.Type ?? "no_action";
            enrichment.SuggestedDraft = parsed.SuggestedAction?.Draft;
            enrichment.SuggestedNotes = parsed.SuggestedAction?.Notes;
            enrichment.Reasoning = parsed.Reasoning;
            enrichment.FlaggedQuestion = parsed.FlagQuestion;
            enrichment.Model = model;
            enrichment.RawResponse = rawResponse;

            // Wipe stale pending tasks/questions before writing fresh ones so re-enrichment doesn't accumulate duplicates
            await _taskRepository.SoftDeletePendingByTicketIdAsync(ticketId);
            await _questionRepository.SoftDeletePendingByTicketIdAsync(ticketId);

            if (parsed.FlagQuestion && !string.IsNullOrWhiteSpace(parsed.QuestionForMaintainer))
            {
                var question = new MaxQuestion
                {
                    CompanyId = ticket.CompanyId,
                    ProjectId = ticket.ProjectId,
                    SourceTicketId = ticketId,
                    Question = parsed.QuestionForMaintainer!,
                    ContextExcerpt = parsed.QuestionContextExcerpt,
                    Status = "pending",
                };
                var created = await _questionRepository.CreateAsync(question);
                enrichment.QuestionId = created.Id;
            }

            await _enrichmentRepository.UpsertAsync(enrichment);

            await CreateTasksFromParsedAsync(ticket, enrichment, parsed);

            _logger.LogInformation(
                "Max enriched ticket {TicketId}: category={Category} confidence={Confidence:F2} action={Action}",
                ticketId, enrichment.Category, enrichment.Confidence, enrichment.SuggestedActionType);

            return enrichment;
        }
        catch (OperationCanceledException)
        {
            _logger.LogInformation("Enrichment canceled for ticket {TicketId}", ticketId);
            if (ticket != null) await TrySetEnrichmentStatusAsync(ticket, "failed");
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to enrich ticket {TicketId}", ticketId);
            if (ticket != null) await TrySetEnrichmentStatusAsync(ticket, "failed");
            return null;
        }
    }

    private async Task SetEnrichmentStatusAsync(Ticket ticket, string status)
    {
        var existing = await _enrichmentRepository.GetByTicketIdAsync(ticket.Id);
        if (existing == null)
        {
            await _enrichmentRepository.UpsertAsync(new MaxTicketEnrichment
            {
                CompanyId = ticket.CompanyId,
                ProjectId = ticket.ProjectId,
                TicketId = ticket.Id,
                Status = status,
            });
        }
        else
        {
            existing.Status = status;
            await _enrichmentRepository.UpsertAsync(existing);
        }
    }

    private async Task TrySetEnrichmentStatusAsync(Ticket ticket, string status)
    {
        try { await SetEnrichmentStatusAsync(ticket, status); }
        catch (Exception ex) { _logger.LogWarning(ex, "Failed to write enrichment status {Status} for ticket {TicketId}", status, ticket.Id); }
    }

    private async Task CreateTasksFromParsedAsync(Ticket ticket, MaxTicketEnrichment enrichment, ParsedEnrichment parsed)
    {
        await CreateTaskFromActionAsync(ticket, enrichment, parsed.SuggestedAction, isPrimary: true);
        if (parsed.AdditionalActions != null)
        {
            foreach (var action in parsed.AdditionalActions)
            {
                if (action?.Type == parsed.SuggestedAction?.Type) continue;
                await CreateTaskFromActionAsync(ticket, enrichment, action, isPrimary: false);
            }
        }
    }

    private async Task CreateTaskFromActionAsync(Ticket ticket, MaxTicketEnrichment enrichment, ParsedSuggestedAction? action, bool isPrimary)
    {
        var actionType = action?.Type ?? "no_action";
        if (actionType == "no_action") return;

        var details = new MaxTaskDetails();
        switch (actionType)
        {
            case "draft_reply":
                details.Draft = action?.Draft;
                details.Notes = action?.Notes;
                break;
            case "merge_duplicate":
                details.DuplicateOfTicketId = enrichment.DuplicateOfTicketId;
                details.Notes = action?.Notes;
                break;
            case "add_to_backlog":
                details.SuggestedTitle = !string.IsNullOrWhiteSpace(action?.Notes) ? action.Notes : enrichment.Summary;
                details.SuggestedKanbanType = MapCategoryToKanbanType(enrichment.Category);
                details.Notes = action?.Notes;
                break;
            case "link_to_story":
                // Worker doesn't have kanban repos, so we can't look up the card
                // title here. Store the number; the SPA reads card details when
                // rendering the task card.
                if (action?.CardNumber == null) return;
                details.LinkToStoryCardNumber = action.CardNumber;
                details.Notes = action?.Notes;
                break;
            case "escalated":
            case "investigate":
                details.Notes = action?.Notes;
                break;
        }

        if (isPrimary && enrichment.FlaggedQuestion && !string.IsNullOrEmpty(enrichment.QuestionId))
            details.QuestionId = enrichment.QuestionId;

        var task = new MaxTask
        {
            CompanyId = ticket.CompanyId,
            ProjectId = ticket.ProjectId,
            TicketId = ticket.Id,
            Type = actionType,
            Status = "pending",
            Confidence = enrichment.Confidence,
            Details = details,
        };
        await _taskRepository.CreateAsync(task);
    }

    private static string MapCategoryToKanbanType(string category) => category switch
    {
        "bug" => "Bug",
        "feature_request" => "Feature",
        _ => "Improvement",
    };

    private static bool MentionsEmDashes(string? toneAvoid)
    {
        if (string.IsNullOrWhiteSpace(toneAvoid)) return false;
        var lower = toneAvoid.ToLowerInvariant();
        return lower.Contains("em-dash") || lower.Contains("em dash") || lower.Contains("emdash");
    }

    private static string StripEmDashes(string text)
    {
        text = text.Replace(" — ", ", ").Replace(" – ", ", ");
        text = text.Replace("—", ",").Replace("–", ",");
        while (text.Contains(",,")) text = text.Replace(",,", ",");
        text = text.Replace(", ,", ",");
        return text;
    }

    private static string BuildSystemPrompt(MaxSettings max, List<MaxInstruction> instructions, List<MaxExampleReply> examples)
    {
        var sb = new StringBuilder();
        sb.AppendLine("You are Max, the AI assistant for a support helpdesk run by a solo maintainer. Your job in this call is to triage one incoming support message and return a structured JSON object that organizes it for the maintainer.");
        sb.AppendLine();
        sb.AppendLine("You are not replying to the user directly. You are not solving their problem. You are filing the ticket so the maintainer can act on it efficiently when they triage. Optimize for the maintainer's time, not the user's experience.");
        sb.AppendLine();

        sb.AppendLine("## Project context");
        sb.AppendLine("<project_context>");
        sb.AppendLine(string.IsNullOrWhiteSpace(max.ProjectContext) ? "(not provided)" : max.ProjectContext);
        sb.AppendLine("</project_context>");
        sb.AppendLine();

        sb.AppendLine("## How to sound when drafting replies");
        sb.AppendLine("<tone_guide>");
        sb.AppendLine(string.IsNullOrWhiteSpace(max.ToneGuide) ? "(not provided)" : max.ToneGuide);
        sb.AppendLine("</tone_guide>");
        sb.AppendLine();

        sb.AppendLine("## Things you MUST NEVER include in replies");
        sb.AppendLine("These are hard rules from the maintainer. Treat them as absolute constraints. If anything on this list would normally appear in your draft (including patterns you tend to default to), rewrite to remove it before returning. Do not include forbidden patterns in code blocks, quoted material, or anywhere else in the reply text.");
        sb.AppendLine("<tone_avoid>");
        sb.AppendLine(string.IsNullOrWhiteSpace(max.ToneAvoid) ? "(not provided)" : max.ToneAvoid);
        sb.AppendLine("</tone_avoid>");
        sb.AppendLine();

        sb.AppendLine("## Example replies in the maintainer's voice");
        sb.AppendLine("<example_replies>");
        if (examples.Count == 0)
        {
            sb.AppendLine("(none provided)");
        }
        else
        {
            var capped = examples.Take(MaxExampleRepliesInPrompt).ToList();
            for (int i = 0; i < capped.Count; i++)
            {
                if (i > 0) sb.AppendLine("---");
                sb.AppendLine(capped[i].ReplyText);
            }
        }
        sb.AppendLine("</example_replies>");
        sb.AppendLine();

        sb.AppendLine("## Specific instructions you have been given");
        sb.AppendLine("<instructions>");
        var relevantInstructions = instructions
            .Where(i => i.Contexts.Contains("enrichment") || i.Contexts.Contains("all") || i.Contexts.Contains("reply"))
            .ToList();
        if (relevantInstructions.Count == 0)
        {
            sb.AppendLine("(none provided)");
        }
        else
        {
            foreach (var inst in relevantInstructions)
                sb.AppendLine($"- {inst.Instruction}");
        }
        sb.AppendLine("</instructions>");
        sb.AppendLine();

        sb.AppendLine("## Categorization");
        sb.AppendLine("Pick exactly one category from: how_to, bug, feature_request, billing, account, feedback, spam, unsure.");
        sb.AppendLine("- how_to: user is confused about how to use a feature that exists");
        sb.AppendLine("- bug: something is not working as designed");
        sb.AppendLine("- feature_request: user wants something the product doesn't do");
        sb.AppendLine("- billing: subscriptions, charges, refunds, card issues");
        sb.AppendLine("- account: login, password reset, account deletion, data export");
        sb.AppendLine("- feedback: opinion, thanks, complaint with no specific action");
        sb.AppendLine("- spam: promotional, off-topic, automated");
        sb.AppendLine("- unsure: does not cleanly fit or message is too vague");
        sb.AppendLine();

        sb.AppendLine("## Summary");
        sb.AppendLine("One sentence, max 140 chars, written for the maintainer. State the issue, not the user's framing. Drop pleasantries.");
        sb.AppendLine();

        sb.AppendLine("## Duplicate detection vs linking");
        sb.AppendLine("Set `duplicate_of` ONLY when the SAME customer (same email address) opened multiple tickets about the same issue. A merge collapses one ticket into another; that only makes sense when there's one conversation that got split.");
        sb.AppendLine();
        sb.AppendLine("When DIFFERENT customers report the same underlying issue:");
        sb.AppendLine("- Add the related ticket ids to `related_ids` to surface the connection");
        sb.AppendLine("- Leave `duplicate_of` null");
        sb.AppendLine("- Pick `draft_reply` (or `investigate` for bugs) as the suggested_action — each customer still needs their own reply");
        sb.AppendLine("- In the suggested_action.notes, mention the pattern (e.g., \"3 customers have reported this in the past week — consider adding to backlog\")");
        sb.AppendLine();
        sb.AppendLine("Never use `merge_duplicate` across different customer emails. Merging would lose one customer's conversation.");
        sb.AppendLine("When unsure between duplicate and related, default to related — false duplicates destroy real signal.");
        sb.AppendLine();

        sb.AppendLine("## Suggested action types");
        sb.AppendLine("- draft_reply: you can write a useful response. Include the draft.");
        sb.AppendLine("- investigate: bug needing maintainer's eyes on code. Include investigation hints in notes.");
        sb.AppendLine("- merge_duplicate: `duplicate_of` is set AND the customer emails match. Send a thanks/tracked-here ack.");
        sb.AppendLine("- add_to_backlog: feature request worth tracking, OR a recurring bug pattern that has NO existing kanban story. Include a kanban-suitable title in notes.");
        sb.AppendLine("- link_to_story: an existing non-Done kanban story in `<available_stories>` already covers this issue. Set `card_number` to the matching story's cardNumber. Prefer this over add_to_backlog whenever a matching story exists.");
        sb.AppendLine("- no_action: spam, off-topic, or feedback that needs no response.");
        sb.AppendLine("- escalated: ambiguous, emotionally charged, legal/safety/refund disputes, or confidence below 0.5.");
        sb.AppendLine();

        sb.AppendLine("## Additional actions");
        sb.AppendLine("A single ticket can warrant more than one thing. The PRIMARY suggested_action is the most user-facing thing (usually replying to the customer). In `additional_actions`, you can include extra actions the maintainer should also consider — currently this is most useful for `add_to_backlog` when a bug or feature request also deserves a tracking story.");
        sb.AppendLine();
        sb.AppendLine("Use additional_actions when:");
        sb.AppendLine("- Primary action is draft_reply or investigate AND the ticket describes a bug that needs a code fix or a feature request worth tracking. Pick ONE of:");
        sb.AppendLine("  - `link_to_story` with `card_number` set to a matching non-Done story in `<available_stories>` (preferred if a match exists)");
        sb.AppendLine("  - `add_to_backlog` with a kanban-suitable title in notes (only if no matching open story exists)");
        sb.AppendLine("- Don't duplicate the primary action in additional_actions");
        sb.AppendLine("- If primary is already add_to_backlog, link_to_story, merge_duplicate, no_action, or escalated, leave additional_actions empty");
        sb.AppendLine();

        sb.AppendLine("## Use existing kanban stories before creating new ones");
        sb.AppendLine("Three sources of kanban context matter:");
        sb.AppendLine("1. `<existing_links>` — kanban stories ALREADY linked to the current ticket.");
        sb.AppendLine("2. Each `<open_issues>` item has `linked_stories` — stories already linked to OTHER recent tickets.");
        sb.AppendLine("3. `<available_stories>` — every non-Done kanban story in the project, whether or not it's linked to anything.");
        sb.AppendLine();
        sb.AppendLine("Decision order when the ticket describes a bug or feature request:");
        sb.AppendLine("a) If the current ticket already has a non-Done story in `existing_links` that covers it → don't suggest anything new. Mention the existing story in reasoning.");
        sb.AppendLine("b) Otherwise, scan `available_stories` for a non-Done story whose title/type clearly matches this ticket's issue → suggest `link_to_story` with that story's `card_number`. This is the most common case after the project has been running for a while.");
        sb.AppendLine("c) Only if no matching story exists anywhere → suggest `add_to_backlog` with a kanban-suitable title.");
        sb.AppendLine();
        sb.AppendLine("Other rules:");
        sb.AppendLine("- You MAY suggest `add_to_backlog` if all matching stories have `is_done: true` and the issue has resurfaced — note that in your reasoning.");
        sb.AppendLine("- Do NOT include ticket ids in `related_ids` if they are already in `linked_ticket_ids`.");
        sb.AppendLine("- When you mention an existing story in your reasoning or notes, use the `STR-N` format.");
        sb.AppendLine("- Don't enumerate candidate stories you considered and rejected. Only mention stories you're recommending action on (or that are already linked). The maintainer doesn't need to see your scratch work.");
        sb.AppendLine();

        sb.AppendLine("## Reply drafting");
        sb.AppendLine("Follow the tone guide, tone avoid, examples, and instructions. Never invent features, fixes, timelines, or refund amounts. Never apologize for bugs — say \"I'll investigate\" or \"tracked here\" instead. 4 sentences or fewer unless genuinely required.");
        sb.AppendLine();

        sb.AppendLine("## Confidence");
        sb.AppendLine("0 to 1. Be honest. Below 0.5 should set action to escalated.");
        sb.AppendLine();

        sb.AppendLine("## Flagging questions");
        sb.AppendLine("If the ticket references a feature/concept/term NOT in the project context, and you cannot confidently respond without it, set flag_question=true and fill in question_for_maintainer and question_context_excerpt. Still produce your best-guess category/summary/action. Don't flag for things obvious from the context — re-read first.");
        sb.AppendLine();

        sb.AppendLine("## Final check before output");
        sb.AppendLine("Re-read your drafted reply (suggested_action.draft, if you wrote one). For every item in <tone_avoid>, verify the draft does not contain it. If anything slipped through, rewrite the draft to remove it before producing the JSON. This is not optional.");
        sb.AppendLine();

        sb.AppendLine("## Output schema");
        sb.AppendLine("Return JSON matching exactly this shape. No preamble, no code fences, no trailing prose.");
        sb.AppendLine();
        sb.AppendLine("{");
        sb.AppendLine("  \"category\": \"how_to | bug | feature_request | billing | account | feedback | spam | unsure\",");
        sb.AppendLine("  \"summary\": \"string, max 140 chars\",");
        sb.AppendLine("  \"confidence\": 0.0,");
        sb.AppendLine("  \"duplicate_of\": null,");
        sb.AppendLine("  \"related_ids\": [],");
        sb.AppendLine("  \"platform\": \"ios | android | web | unknown\",");
        sb.AppendLine("  \"tags\": [],");
        sb.AppendLine("  \"suggested_action\": {");
        sb.AppendLine("    \"type\": \"draft_reply | investigate | merge_duplicate | add_to_backlog | link_to_story | no_action | escalated\",");
        sb.AppendLine("    \"draft\": null,");
        sb.AppendLine("    \"notes\": null,");
        sb.AppendLine("    \"card_number\": null");
        sb.AppendLine("  },");
        sb.AppendLine("  \"additional_actions\": [");
        sb.AppendLine("    { \"type\": \"link_to_story\", \"card_number\": 6, \"notes\": \"Story STR-6 already tracks this bug\" }");
        sb.AppendLine("  ],");
        sb.AppendLine("  \"flag_question\": false,");
        sb.AppendLine("  \"question_for_maintainer\": null,");
        sb.AppendLine("  \"question_context_excerpt\": null,");
        sb.AppendLine("  \"reasoning\": \"one sentence explaining the category choice\"");
        sb.AppendLine("}");

        return sb.ToString();
    }

    private static string BuildUserMessage(Ticket ticket, TicketMessage? latestCustomerMessage, List<Ticket> openTickets, Dictionary<string, MaxTicketEnrichment> enrichmentByTicketId)
    {
        var openIssuesPayload = openTickets.Select(t =>
        {
            enrichmentByTicketId.TryGetValue(t.Id, out var e);
            return new
            {
                id = t.Id,
                ticketNumber = t.TicketNumber,
                subject = t.Subject,
                customerEmail = t.CustomerEmail,
                category = e?.Category,
                summary = e?.Summary,
                tags = e?.Tags ?? new List<string>(),
                // Worker doesn't have kanban repos. The API enrichment path (manual
                // re-enrich, intake/ticket creation) populates this; here we emit
                // an empty array so the prompt schema stays consistent.
                linked_stories = new List<object>(),
            };
        }).ToList();

        // Worker only enriches brand-new tickets that don't have existing links
        // yet. We still emit the block so the system prompt's existing-links
        // rules apply consistently across API and worker enrichment.
        var existingLinksPayload = new
        {
            linked_ticket_ids = new List<string>(),
            linked_stories = new List<object>(),
        };

        var sb = new StringBuilder();
        sb.AppendLine("<open_issues>");
        sb.AppendLine(JsonSerializer.Serialize(openIssuesPayload, new JsonSerializerOptions { WriteIndented = false }));
        sb.AppendLine("</open_issues>");
        sb.AppendLine();
        sb.AppendLine("<existing_links>");
        sb.AppendLine(JsonSerializer.Serialize(existingLinksPayload, new JsonSerializerOptions { WriteIndented = false }));
        sb.AppendLine("</existing_links>");
        sb.AppendLine();
        // Worker has no kanban repos. Maintainer-driven manual re-enrich uses
        // the API path, which populates the real list; we send empty here so
        // the prompt schema stays consistent across both code paths.
        sb.AppendLine("<available_stories>");
        sb.AppendLine("[]");
        sb.AppendLine("</available_stories>");
        sb.AppendLine();
        sb.AppendLine("<ticket>");
        sb.AppendLine($"Source: {ticket.CreationSource}");
        sb.AppendLine($"From: {ticket.CustomerName ?? "(unknown)"} <{ticket.CustomerEmail ?? "(unknown)"}>");
        sb.AppendLine($"Subject: {ticket.Subject}");
        sb.AppendLine("Body:");
        sb.AppendLine(latestCustomerMessage?.Body ?? "(no customer message body found)");
        sb.AppendLine("</ticket>");

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
            _logger.LogWarning("Anthropic returned {Status} for enrichment: {Body}", (int)response.StatusCode, body);
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(body);
            if (!doc.RootElement.TryGetProperty("content", out var content) || content.GetArrayLength() == 0)
            {
                _logger.LogWarning("Anthropic response missing content array");
                return null;
            }
            var first = content[0];
            if (!first.TryGetProperty("text", out var text)) return null;
            return text.GetString();
        }
        catch (JsonException ex)
        {
            _logger.LogWarning(ex, "Failed to parse Anthropic envelope");
            return null;
        }
    }

    private static ParsedEnrichment? ParseEnrichmentResponse(string json)
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

            return JsonSerializer.Deserialize<ParsedEnrichment>(trimmed, new JsonSerializerOptions
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

    private class ParsedEnrichment
    {
        public string? Category { get; set; }
        public string? Summary { get; set; }
        public double Confidence { get; set; }
        public string? DuplicateOf { get; set; }
        public List<string>? RelatedIds { get; set; }
        public string? Platform { get; set; }
        public List<string>? Tags { get; set; }
        public ParsedSuggestedAction? SuggestedAction { get; set; }
        public List<ParsedSuggestedAction>? AdditionalActions { get; set; }
        public bool FlagQuestion { get; set; }
        public string? QuestionForMaintainer { get; set; }
        public string? QuestionContextExcerpt { get; set; }
        public string? Reasoning { get; set; }
    }

    private class ParsedSuggestedAction
    {
        public string? Type { get; set; }
        public string? Draft { get; set; }
        public string? Notes { get; set; }
        public long? CardNumber { get; set; }
    }
}
