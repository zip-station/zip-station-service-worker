using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using MimeKit;
using MongoDB.Driver;
using ZipStation.Worker.Entities;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Repositories;

namespace ZipStation.Worker.Services;

public interface IEmailPollingService
{
    Task PollAllProjectsAsync(CancellationToken ct);
    Task ImportHistoryAsync(CancellationToken ct);
}

public class EmailPollingService : IEmailPollingService
{
    private readonly ILogger<EmailPollingService> _logger;
    private readonly ProjectRepository _projectRepository;
    private readonly IntakeEmailRepository _intakeEmailRepository;
    private readonly IntakeRuleRepository _intakeRuleRepository;
    private readonly TicketRepository _ticketRepository;
    private readonly TicketMessageRepository _ticketMessageRepository;
    private readonly CustomerRepository _customerRepository;
    private readonly TicketIdCounterRepository _ticketIdCounterRepository;
    private readonly MongoDB.Driver.IMongoDatabase _database;
    private readonly Helpers.AppConfig _appConfig;
    private static readonly HttpClient _webhookClient = new() { Timeout = TimeSpan.FromSeconds(10) };

    public EmailPollingService(
        ILogger<EmailPollingService> logger,
        ProjectRepository projectRepository,
        IntakeEmailRepository intakeEmailRepository,
        IntakeRuleRepository intakeRuleRepository,
        TicketRepository ticketRepository,
        TicketMessageRepository ticketMessageRepository,
        CustomerRepository customerRepository,
        TicketIdCounterRepository ticketIdCounterRepository,
        MongoDB.Driver.IMongoDatabase database,
        Microsoft.Extensions.Options.IOptions<Helpers.AppConfig> appConfig)
    {
        _logger = logger;
        _projectRepository = projectRepository;
        _intakeEmailRepository = intakeEmailRepository;
        _intakeRuleRepository = intakeRuleRepository;
        _ticketRepository = ticketRepository;
        _ticketMessageRepository = ticketMessageRepository;
        _customerRepository = customerRepository;
        _ticketIdCounterRepository = ticketIdCounterRepository;
        _database = database;
        _appConfig = appConfig.Value;
    }

    public async Task PollAllProjectsAsync(CancellationToken ct)
    {
        var projects = await _projectRepository.GetAllWithImapAsync();

        if (projects.Count == 0)
        {
            _logger.LogDebug("No projects with IMAP configured, skipping email poll");
            return;
        }

        foreach (var project in projects)
        {
            if (ct.IsCancellationRequested) break;

            try
            {
                await PollProjectAsync(project, ct);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error polling emails for project {ProjectId} ({ProjectName})",
                    project.Id, project.Name);
            }
        }
    }

    private async Task PollProjectAsync(Project project, CancellationToken ct)
    {
        var imap = project.Settings.Imap;
        if (imap == null || string.IsNullOrEmpty(imap.Host))
        {
            _logger.LogDebug("Skipping project {ProjectName} — no IMAP settings", project.Name);
            return;
        }

        _logger.LogInformation("Polling emails for project {ProjectName} via {ImapHost}:{ImapPort} (user: {ImapUser}, ssl: {UseSsl})",
            project.Name, imap.Host, imap.Port, imap.Username, imap.UseSsl);

        using var client = new ImapClient();

        // Accept server certificates (MXRoute shared hosting may have mismatched certs)
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        try
        {
            var secureSocketOptions = imap.UseSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
            await client.ConnectAsync(imap.Host, imap.Port, secureSocketOptions, ct);

            _logger.LogInformation("Connected to {ImapHost}, authenticating as {ImapUser}...", imap.Host, imap.Username);
            var decryptedPassword = EncryptionHelper.Decrypt(imap.Password);
            await client.AuthenticateAsync(imap.Username, decryptedPassword, ct);

            var inbox = client.Inbox;
            await inbox.OpenAsync(FolderAccess.ReadWrite, ct);

            _logger.LogInformation("Connected to {ImapHost} — {MessageCount} messages in inbox",
                imap.Host, inbox.Count);

            if (inbox.Count == 0)
            {
                await client.DisconnectAsync(true, ct);
                return;
            }

            // Fetch all messages (process newest first, but we read all)
            var messages = await inbox.FetchAsync(0, -1,
                MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope | MessageSummaryItems.Flags, ct);

            var processedCount = 0;
            var uidsToDelete = new List<UniqueId>();

            foreach (var summary in messages)
            {
                if (ct.IsCancellationRequested) break;
                if (summary.Flags.HasValue && summary.Flags.Value.HasFlag(MessageFlags.Seen)) continue;

                try
                {
                    var message = await inbox.GetMessageAsync(summary.UniqueId, ct);
                    var created = await ProcessEmailAsync(project, message, ct);

                    if (created)
                    {
                        // Mark as seen
                        await inbox.AddFlagsAsync(summary.UniqueId, MessageFlags.Seen, true, ct);
                        processedCount++;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error processing email UID {UniqueId} in project {ProjectName}",
                        summary.UniqueId, project.Name);
                }
            }

            _logger.LogInformation("Processed {Count} new emails for project {ProjectName}",
                processedCount, project.Name);
        }
        catch (AuthenticationException ex)
        {
            _logger.LogError("IMAP authentication failed for project {ProjectName} (host: {Host}, user: {User}): {Error}",
                project.Name, imap.Host, imap.Username, ex.Message);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogError(ex, "IMAP connection error for project {ProjectName}", project.Name);
        }
        finally
        {
            if (client.IsConnected)
                await client.DisconnectAsync(true, ct);
        }
    }

    private async Task<bool> ProcessEmailAsync(Project project, MimeMessage message, CancellationToken ct)
    {
        var fromAddress = message.From.Mailboxes.FirstOrDefault();
        if (fromAddress == null) return false;

        var messageId = message.MessageId;
        var inReplyTo = message.InReplyTo;

        // Skip if we already processed this message
        if (!string.IsNullOrEmpty(messageId))
        {
            var existing = await _intakeEmailRepository.GetByMessageIdAsync(messageId);
            if (existing != null)
            {
                _logger.LogDebug("Skipping already-processed email {MessageId}", messageId);
                return false;
            }
        }

        // Check if this is a reply to an existing ticket
        // Try threading via In-Reply-To header first, then fall back to matching by customer email
        {
            var handled = await TryThreadToExistingTicketAsync(project, message, fromAddress, ct);
            if (handled) return true;
        }

        // Create intake email — prefer plain text, fall back to stripped HTML
        var bodyText = message.TextBody ?? "";
        var bodyHtml = message.HtmlBody;

        // If no plain text body, extract text from HTML
        if (string.IsNullOrWhiteSpace(bodyText) && !string.IsNullOrEmpty(bodyHtml))
        {
            bodyText = StripHtmlToText(bodyHtml);
        }

        var actualFromEmail = fromAddress.Address;
        var actualFromName = fromAddress.Name ?? fromAddress.Address.Split('@')[0];
        var actualSubject = message.Subject ?? "(no subject)";
        var actualBodyText = bodyText;

        // Check if this is a system-generated contact form email
        var contactForm = project.Settings?.ContactForm;
        _logger.LogInformation("Contact form check: enabled={Enabled}, senders=[{Senders}], from={From}, bodyLen={BodyLen}",
            contactForm?.Enabled, contactForm?.SystemSenderEmails != null ? string.Join(",", contactForm.SystemSenderEmails) : "null", fromAddress.Address, bodyText.Length);

        if (contactForm != null && contactForm.Enabled &&
            contactForm.SystemSenderEmails.Any(e => e.Equals(fromAddress.Address, StringComparison.OrdinalIgnoreCase)))
        {
            _logger.LogInformation("Detected contact form email from system sender {Sender}, parsing body: {BodyPreview}",
                fromAddress.Address, bodyText.Length > 200 ? bodyText[..200] + "..." : bodyText);

            var parsed = ParseContactFormFields(bodyText, contactForm);
            if (!string.IsNullOrEmpty(parsed.Email))
            {
                actualFromEmail = parsed.Email;
                actualFromName = parsed.Name ?? parsed.Email.Split('@')[0];
                actualBodyText = parsed.Message ?? bodyText;
                if (!string.IsNullOrEmpty(parsed.Subject))
                    actualSubject = parsed.Subject;

                _logger.LogInformation("Parsed contact form: email={Email}, name={Name}", actualFromEmail, actualFromName);
            }
        }

        var intake = new IntakeEmail
        {
            CompanyId = project.CompanyId,
            ProjectId = project.Id,
            FromEmail = actualFromEmail,
            FromName = actualFromName,
            Subject = actualSubject,
            BodyText = actualBodyText,
            BodyHtml = bodyHtml,
            ReceivedOn = message.Date.ToUnixTimeMilliseconds(),
            Status = 0, // Pending
            MessageId = messageId,
            InReplyTo = inReplyTo
        };

        // Apply intake rules — check against BOTH original sender and parsed fields
        // so rules like "FromEmail = contact@softgoods.app" still match system-generated emails
        var rules = await _intakeRuleRepository.GetEnabledByProjectIdAsync(project.Id);
        var ruleCheckEmail = new IntakeEmail
        {
            FromEmail = fromAddress.Address, // original raw sender
            Subject = message.Subject ?? "(no subject)", // original subject
            BodyText = bodyText // full body text
        };
        var matchedRule = MatchIntakeRule(ruleCheckEmail, rules);

        if (matchedRule != null)
        {
            _logger.LogInformation("Intake rule '{RuleName}' matched for email from {From} — action: {Action}",
                matchedRule.Name, fromAddress.Address, matchedRule.Action);

            // 0=AutoApprove, 1=AutoDeny, 2=AutoDenyPermanent
            switch (matchedRule.Action)
            {
                case 0: // AutoApprove
                    intake.Status = 1; // Approved
                    intake.ProcessedOn = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                    var created = await _intakeEmailRepository.CreateAsync(intake);
                    await CreateTicketFromIntakeAsync(project, created);
                    return true;

                case 1: // AutoDeny
                    intake.Status = 2; // Denied
                    intake.ProcessedOn = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                    await _intakeEmailRepository.CreateAsync(intake);
                    return true;

                case 2: // AutoDenyPermanent
                    intake.Status = 2; // Denied
                    intake.DeniedPermanently = true;
                    intake.ProcessedOn = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                    await _intakeEmailRepository.CreateAsync(intake);
                    return true;
            }
        }

        // Spam scoring with configurable thresholds
        intake.SpamScore = CalculateSpamScore(intake);

        var spamSettings = project.Settings?.Spam;
        if (spamSettings != null && spamSettings.AutoDenyEnabled && intake.SpamScore >= spamSettings.AutoDenyThreshold)
        {
            intake.Status = 2; // Denied
            intake.ProcessedOn = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
            await _intakeEmailRepository.CreateAsync(intake);
            _logger.LogInformation("Spam auto-denied: {From} (score: {SpamScore}, threshold: {Threshold})",
                fromAddress.Address, intake.SpamScore, spamSettings.AutoDenyThreshold);
            return true;
        }

        await _intakeEmailRepository.CreateAsync(intake);
        _logger.LogInformation("New intake email from {From}: {Subject} (spam: {SpamScore})",
            fromAddress.Address, intake.Subject, intake.SpamScore);

        return true;
    }

    private async Task<bool> TryThreadToExistingTicketAsync(
        Project project, MimeMessage message, MailboxAddress fromAddress, CancellationToken ct)
    {
        // Skip if sender is a system sender (contact form emails should go through parsing)
        var contactForm = project.Settings?.ContactForm;
        if (contactForm != null && contactForm.Enabled &&
            contactForm.SystemSenderEmails.Any(e => e.Equals(fromAddress.Address, StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        Ticket? existingTicket = null;

        // Strategy 1: Check In-Reply-To and References headers for our ticket anchor
        // Our outgoing emails set InReplyTo/References to "<ticket-{ticketId}@domain>"
        // When customers reply, their client preserves these in the References chain
        var ticketIdFromHeaders = TryExtractTicketIdFromHeaders(message);
        if (ticketIdFromHeaders != null)
        {
            _logger.LogInformation("Threading check: found ticket ID {TicketId} in email headers", ticketIdFromHeaders);
            existingTicket = await _ticketRepository.GetByIdAsync(ticketIdFromHeaders);
            if (existingTicket != null)
                _logger.LogInformation("Found ticket {TicketId} (status={Status}) by email headers",
                    existingTicket.Id, existingTicket.Status);
        }

        // Strategy 2: Try to extract ticket ID from subject line
        // Subject looks like "Re: ProjectName - Ticket SG-001" or "RE: Fwd: ProjectName - Ticket 001"
        if (existingTicket == null)
        {
            var subject = message.Subject ?? "";
            var ticketNumber = TryParseTicketNumberFromSubject(subject, project.Settings?.TicketId);
            if (ticketNumber.HasValue)
            {
                _logger.LogInformation("Threading check: parsed ticket number {TicketNumber} from subject \"{Subject}\"",
                    ticketNumber.Value, subject);
                existingTicket = await _ticketRepository.GetByTicketNumberAndProjectAsync(ticketNumber.Value, project.Id);
                if (existingTicket != null)
                    _logger.LogInformation("Found ticket {TicketId} (status={Status}) by ticket number {TicketNumber}",
                        existingTicket.Id, existingTicket.Status, ticketNumber.Value);
            }
        }

        // Strategy 3: Fall back to matching by customer email (Open/Pending only)
        if (existingTicket == null)
        {
            _logger.LogInformation("Threading check: looking for open/pending ticket from {Email} in project {ProjectId}",
                fromAddress.Address, project.Id);
            existingTicket = await _ticketRepository.GetByCustomerEmailAndProjectAsync(
                fromAddress.Address, project.Id);
        }

        if (existingTicket == null)
        {
            _logger.LogInformation("No existing ticket found for threading (email={Email}, subject=\"{Subject}\")",
                fromAddress.Address, message.Subject ?? "");
            return false;
        }

        // If ticket was Closed or Resolved, reopen it
        if (existingTicket.Status is 2 or 3) // Resolved=2, Closed=3
        {
            _logger.LogInformation("Reopening ticket {TicketId} (was status={Status}) due to customer reply",
                existingTicket.Id, existingTicket.Status);
            await _ticketRepository.UpdateStatusAsync(existingTicket.Id, 0); // Open
        }

        var bodyText = message.TextBody ?? "";
        var bodyHtml = message.HtmlBody;

        // Add as a new message on the existing ticket
        var ticketMessage = new TicketMessage
        {
            TicketId = existingTicket.Id,
            CompanyId = project.CompanyId,
            ProjectId = project.Id,
            Body = bodyText,
            BodyHtml = bodyHtml,
            IsInternalNote = false,
            AuthorName = fromAddress.Name ?? fromAddress.Address.Split('@')[0],
            AuthorEmail = fromAddress.Address,
            Source = 0 // Customer
        };

        await _ticketMessageRepository.CreateAsync(ticketMessage);

        _logger.LogInformation("Threaded reply from {From} to ticket {TicketId}",
            fromAddress.Address, existingTicket.Id);

        return true;
    }

    /// <summary>
    /// Extracts a ticket ID from email In-Reply-To and References headers.
    /// Our outgoing emails use the pattern: &lt;ticket-{mongoId}@domain&gt;
    /// and &lt;{messageId}@domain&gt;. The customer's reply will reference these.
    /// </summary>
    private static string? TryExtractTicketIdFromHeaders(MimeMessage message)
    {
        // Collect all header values to check: In-Reply-To + References
        var headerValues = new List<string>();
        if (!string.IsNullOrEmpty(message.InReplyTo))
            headerValues.Add(message.InReplyTo);
        foreach (var reference in message.References)
            headerValues.Add(reference);

        // Look for our ticket anchor pattern: ticket-{24-char hex ObjectId}@
        var pattern = @"ticket-([a-f0-9]{24})@";
        foreach (var header in headerValues)
        {
            var match = System.Text.RegularExpressions.Regex.Match(header, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (match.Success)
                return match.Groups[1].Value;
        }

        return null;
    }

    /// <summary>
    /// Tries to extract a ticket number from an email subject line.
    /// Handles subjects like "Re: Softgoods - Ticket SG-001", "RE: FW: Softgoods - Ticket 042", etc.
    /// Matches the project's configured prefix if set.
    /// </summary>
    private static long? TryParseTicketNumberFromSubject(string subject, TicketIdSettings? settings)
    {
        if (string.IsNullOrEmpty(subject)) return null;

        var prefix = settings?.Prefix ?? "";

        // Build a regex pattern to match the ticket ID in the subject
        // With prefix: "SG-001", "SG-42", "SG-0001"
        // Without prefix: just digits like "001", "42"
        string pattern;
        if (!string.IsNullOrEmpty(prefix))
        {
            // Match prefix-digits (case insensitive), e.g. SG-001
            pattern = $@"(?:^|[\s\-])({System.Text.RegularExpressions.Regex.Escape(prefix)}-(\d+))(?:\s|$|[)\]])";
        }
        else
        {
            // Without prefix, look for "Ticket NNN" pattern to avoid matching random numbers
            pattern = @"[Tt]icket\s+(\d+)";
        }

        var match = System.Text.RegularExpressions.Regex.Match(subject, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (!match.Success) return null;

        // Extract the numeric part
        var numberStr = !string.IsNullOrEmpty(prefix) ? match.Groups[2].Value : match.Groups[1].Value;
        if (long.TryParse(numberStr, out var ticketNumber) && ticketNumber > 0)
            return ticketNumber;

        return null;
    }

    private async Task CreateTicketFromIntakeAsync(Project project, IntakeEmail intake)
    {
        // Get or create customer
        var customer = await _customerRepository.GetByEmailAndProjectAsync(intake.FromEmail, project.Id);
        if (customer == null)
        {
            customer = new Customer
            {
                CompanyId = project.CompanyId,
                ProjectId = project.Id,
                Email = intake.FromEmail,
                Name = intake.FromName
            };
            await _customerRepository.CreateAsync(customer);
        }

        // Generate ticket number
        var ticketIdSettings = project.Settings.TicketId;
        var ticketNumber = await GenerateTicketNumberAsync(project.Id, ticketIdSettings);
        var displayId = FormatTicketId(ticketNumber, ticketIdSettings);
        var subject = (ticketIdSettings.SubjectTemplate ?? "{ProjectName} - Ticket {TicketId}")
            .Replace("{ProjectName}", project.Name)
            .Replace("{TicketId}", displayId)
            .Replace("{TicketNumber}", ticketNumber.ToString());

        // Always use the template-generated subject — it includes the ticket ID which is
        // needed for threading. The original email subject is preserved in intake.Subject.
        var ticketSubject = subject;

        var ticket = new Ticket
        {
            CompanyId = project.CompanyId,
            ProjectId = project.Id,
            TicketNumber = ticketNumber,
            Subject = ticketSubject,
            Status = 0, // Open
            Priority = 1, // Normal
            CustomerName = intake.FromName,
            CustomerEmail = intake.FromEmail,
            CreationSource = 2 // IntakeAutoRule
        };

        var createdTicket = await _ticketRepository.CreateAsync(ticket);

        // Create first message from customer
        var message = new TicketMessage
        {
            TicketId = createdTicket.Id,
            CompanyId = project.CompanyId,
            ProjectId = project.Id,
            Body = intake.BodyText,
            BodyHtml = intake.BodyHtml,
            IsInternalNote = false,
            AuthorName = intake.FromName,
            AuthorEmail = intake.FromEmail,
            Source = 0 // Customer
        };
        await _ticketMessageRepository.CreateAsync(message);

        // Update intake with ticket reference
        intake.TicketId = createdTicket.Id;
        await _intakeEmailRepository.UpdateAsync(intake);

        // Update customer counts
        customer.TotalTicketCount++;
        customer.OpenTicketCount++;
        await _customerRepository.UpdateAsync(customer);

        // Fire alerts for new ticket
        await FireAlertsForNewTicketAsync(project, createdTicket);

        _logger.LogInformation("Auto-approved intake created ticket {TicketId} for {Email}",
            createdTicket.Id, intake.FromEmail);
    }

    private static IntakeRule? MatchIntakeRule(IntakeEmail intake, List<IntakeRule> rules)
    {
        foreach (var rule in rules)
        {
            if (rule.Conditions.Count == 0) continue;

            // ALL conditions must match (AND logic)
            var allMatch = rule.Conditions.All(condition =>
            {
                // 0=FromEmail, 1=FromDomain, 2=SubjectContains, 3=BodyContains
                return condition.Type switch
                {
                    0 => intake.FromEmail.Equals(condition.Value, StringComparison.OrdinalIgnoreCase),
                    1 => intake.FromEmail.EndsWith($"@{condition.Value.TrimStart('@')}", StringComparison.OrdinalIgnoreCase),
                    2 => intake.Subject.Contains(condition.Value, StringComparison.OrdinalIgnoreCase),
                    3 => intake.BodyText.Contains(condition.Value, StringComparison.OrdinalIgnoreCase),
                    _ => false
                };
            });

            if (allMatch) return rule;
        }
        return null;
    }

    private static int CalculateSpamScore(IntakeEmail intake)
    {
        var score = 0;

        // No subject
        if (string.IsNullOrWhiteSpace(intake.Subject) || intake.Subject == "(no subject)")
            score += 15;

        // Very short body
        if (intake.BodyText.Length < 10)
            score += 10;

        // All caps subject
        if (intake.Subject == intake.Subject.ToUpperInvariant() && intake.Subject.Length > 5)
            score += 20;

        // Common spam indicators
        var spamKeywords = new[] { "unsubscribe", "click here", "act now", "limited time", "winner", "congratulations", "free money" };
        foreach (var keyword in spamKeywords)
        {
            if (intake.BodyText.Contains(keyword, StringComparison.OrdinalIgnoreCase) ||
                intake.Subject.Contains(keyword, StringComparison.OrdinalIgnoreCase))
            {
                score += 15;
            }
        }

        // Excessive links in body
        var linkCount = intake.BodyText.Split("http", StringSplitOptions.None).Length - 1;
        if (linkCount > 5) score += 20;

        return Math.Min(score, 100);
    }

    private async Task<long> GenerateTicketNumberAsync(string projectId, TicketIdSettings settings)
    {
        if (!settings.UseRandomNumbers)
        {
            var next = await _ticketIdCounterRepository.GetNextTicketNumberAsync(projectId);
            return Math.Max(next, settings.StartingNumber > 0 ? settings.StartingNumber : next);
        }

        var min = settings.StartingNumber > 0 ? settings.StartingNumber : (long)Math.Pow(10, Math.Max(settings.MinLength, 3) - 1);
        var max = (long)Math.Pow(10, settings.MaxLength) - 1;
        if (min > max) min = max / 2;

        var random = new Random();
        for (var attempt = 0; attempt < 100; attempt++)
        {
            var candidate = (long)(random.NextDouble() * (max - min + 1)) + min;
            var exists = await _ticketRepository.ExistsByTicketNumberAndProjectAsync(projectId, candidate);
            if (!exists) return candidate;
        }

        return await _ticketIdCounterRepository.GetNextTicketNumberAsync(projectId);
    }

    private static string FormatTicketId(long ticketNumber, TicketIdSettings settings)
    {
        var prefix = settings.Prefix ?? "";
        var minLen = Math.Max(settings.MinLength, 3);
        var idPart = ticketNumber.ToString().PadLeft(minLen, '0');
        return string.IsNullOrEmpty(prefix) ? idPart : $"{prefix}-{idPart}";
    }

    private record ContactFormParsed(string? Email, string? Name, string? Message, string? Subject);

    private static ContactFormParsed ParseContactFormFields(string bodyText, ContactFormSettings settings)
    {
        string? email = null, name = null, message = null, subject = null;

        var lines = bodyText.Split('\n');

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();

            // Try "Label: Value" format (most common)
            var colonIdx = line.IndexOf(':');
            if (colonIdx <= 0) continue;

            var label = line[..colonIdx].Trim();
            var value = line[(colonIdx + 1)..].Trim();

            if (label.Equals(settings.EmailLabel, StringComparison.OrdinalIgnoreCase))
                email = value;
            else if (label.Equals(settings.NameLabel, StringComparison.OrdinalIgnoreCase))
                name = value;
            else if (label.Equals(settings.MessageLabel, StringComparison.OrdinalIgnoreCase))
            {
                // Message might span multiple lines — grab everything after "Message:" until the next known label or end
                var messageLines = new List<string>();
                if (!string.IsNullOrEmpty(value))
                    messageLines.Add(value);

                for (var j = i + 1; j < lines.Length; j++)
                {
                    var nextLine = lines[j].Trim();
                    var nextColon = nextLine.IndexOf(':');
                    // Stop if we hit another labeled field
                    if (nextColon > 0)
                    {
                        var nextLabel = nextLine[..nextColon].Trim();
                        if (nextLabel.Equals(settings.EmailLabel, StringComparison.OrdinalIgnoreCase) ||
                            nextLabel.Equals(settings.NameLabel, StringComparison.OrdinalIgnoreCase) ||
                            (settings.SubjectLabel != null && nextLabel.Equals(settings.SubjectLabel, StringComparison.OrdinalIgnoreCase)))
                            break;
                    }
                    messageLines.Add(nextLine);
                }
                message = string.Join("\n", messageLines).Trim();
            }
            else if (settings.SubjectLabel != null && label.Equals(settings.SubjectLabel, StringComparison.OrdinalIgnoreCase))
                subject = value;
        }

        return new ContactFormParsed(email, name, message, subject);
    }

    public async Task ImportHistoryAsync(CancellationToken ct)
    {
        var projects = await _projectRepository.GetAllWithImapAsync();

        if (projects.Count == 0)
        {
            _logger.LogDebug("No projects with IMAP configured, skipping history import");
            return;
        }

        foreach (var project in projects)
        {
            if (ct.IsCancellationRequested) break;

            try
            {
                await ImportProjectHistoryAsync(project, ct);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing history for project {ProjectId} ({ProjectName})",
                    project.Id, project.Name);
            }
        }
    }

    private async Task ImportProjectHistoryAsync(Project project, CancellationToken ct)
    {
        var imap = project.Settings.Imap;
        if (imap == null || string.IsNullOrEmpty(imap.Host)) return;

        _logger.LogInformation("Importing email history for project {ProjectName} via {ImapHost}",
            project.Name, imap.Host);

        using var client = new ImapClient();
        client.ServerCertificateValidationCallback = (s, c, h, e) => true;

        try
        {
            var secureSocketOptions = imap.UseSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
            await client.ConnectAsync(imap.Host, imap.Port, secureSocketOptions, ct);

            var decryptedPassword = EncryptionHelper.Decrypt(imap.Password);
            await client.AuthenticateAsync(imap.Username, decryptedPassword, ct);

            var inbox = client.Inbox;
            await inbox.OpenAsync(FolderAccess.ReadOnly, ct);

            _logger.LogInformation("History import: {MessageCount} messages in inbox for {ProjectName}",
                inbox.Count, project.Name);

            if (inbox.Count == 0)
            {
                await client.DisconnectAsync(true, ct);
                return;
            }

            var messages = await inbox.FetchAsync(0, -1,
                MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope | MessageSummaryItems.Flags, ct);

            var importedCount = 0;

            foreach (var summary in messages)
            {
                if (ct.IsCancellationRequested) break;

                try
                {
                    var message = await inbox.GetMessageAsync(summary.UniqueId, ct);
                    var messageId = message.MessageId;

                    // Skip if no message ID or already imported
                    if (!string.IsNullOrEmpty(messageId))
                    {
                        var existing = await _intakeEmailRepository.GetByMessageIdAsync(messageId);
                        if (existing != null)
                        {
                            continue; // Already imported
                        }
                    }

                    var created = await ProcessEmailAsync(project, message, ct);
                    if (created) importedCount++;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error importing historical email UID {UniqueId} in project {ProjectName}",
                        summary.UniqueId, project.Name);
                }
            }

            _logger.LogInformation("History import complete: imported {Count} emails for project {ProjectName}",
                importedCount, project.Name);
        }
        catch (AuthenticationException ex)
        {
            _logger.LogError("IMAP authentication failed during history import for project {ProjectName}: {Error}",
                project.Name, ex.Message);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            _logger.LogError(ex, "IMAP connection error during history import for project {ProjectName}", project.Name);
        }
        finally
        {
            if (client.IsConnected)
                await client.DisconnectAsync(true, ct);
        }
    }

    private static string StripHtmlToText(string html)
    {
        // Replace <br>, <br/>, </p>, </div>, </li> with newlines
        var text = System.Text.RegularExpressions.Regex.Replace(html, @"<br\s*/?>|</p>|</div>|</li>|</tr>", "\n", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        // Strip remaining HTML tags
        text = System.Text.RegularExpressions.Regex.Replace(text, @"<[^>]+>", "");
        // Decode HTML entities
        text = System.Net.WebUtility.HtmlDecode(text);
        // Normalize whitespace per line
        var lines = text.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0);
        return string.Join("\n", lines);
    }

    private async Task FireAlertsForNewTicketAsync(Entities.Project project, Entities.Ticket ticket)
    {
        try
        {
            var alertsCollection = _database.GetCollection<Entities.Alert>(_appConfig.ZipStationMongoDb.Collections.Alerts);
            var filter = MongoDB.Driver.Builders<Entities.Alert>.Filter.Eq(a => a.ProjectId, project.Id)
                       & MongoDB.Driver.Builders<Entities.Alert>.Filter.Eq(a => a.TriggerType, 0) // NewTicket
                       & MongoDB.Driver.Builders<Entities.Alert>.Filter.Eq(a => a.IsEnabled, true)
                       & MongoDB.Driver.Builders<Entities.Alert>.Filter.Eq(a => a.IsVoid, false);
            var alerts = await alertsCollection.Find(filter).ToListAsync();

            foreach (var alert in alerts)
            {
                try
                {
                    var message = $"[{project.Name}] New ticket from {ticket.CustomerEmail ?? "unknown"}: {ticket.Subject}";
                    string payload;

                    switch (alert.ChannelType)
                    {
                        case 0: // Slack
                            payload = System.Text.Json.JsonSerializer.Serialize(new { text = message });
                            break;
                        case 1: // Discord
                            payload = System.Text.Json.JsonSerializer.Serialize(new { content = message });
                            break;
                        default: // Generic
                            payload = System.Text.Json.JsonSerializer.Serialize(new
                            {
                                @event = "new_ticket",
                                ticketId = ticket.Id,
                                subject = ticket.Subject,
                                customerEmail = ticket.CustomerEmail ?? "",
                                projectName = project.Name
                            });
                            break;
                    }

                    var content = new System.Net.Http.StringContent(payload, System.Text.Encoding.UTF8, "application/json");
                    var response = await _webhookClient.PostAsync(alert.WebhookUrl, content);

                    if (response.IsSuccessStatusCode)
                        _logger.LogInformation("Worker alert fired: {AlertId} for ticket {TicketId}", alert.Id, ticket.Id);
                    else
                        _logger.LogWarning("Worker alert webhook returned {StatusCode} for {AlertId}", response.StatusCode, alert.Id);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Worker alert failed for {AlertId}", alert.Id);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error firing worker alerts for project {ProjectId}", project.Id);
        }
    }
}
