using MongoDB.Bson;
using MongoDB.Driver;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Services;

namespace ZipStation.Worker;

public class Worker : BackgroundService
{
    private readonly ILogger<Worker> _logger;
    private readonly IEmailPollingService _emailPollingService;
    private readonly IDiscordPollingService _discordPollingService;
    private readonly IMongoDatabase _database;
    private readonly AppConfig _appConfig;

    public Worker(
        ILogger<Worker> logger,
        IEmailPollingService emailPollingService,
        IDiscordPollingService discordPollingService,
        IMongoDatabase database,
        Microsoft.Extensions.Options.IOptions<AppConfig> appConfig)
    {
        _logger = logger;
        _emailPollingService = emailPollingService;
        _discordPollingService = discordPollingService;
        _database = database;
        _appConfig = appConfig.Value;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Zip Station Worker starting");

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                var processes = new List<Task>
                {
                    PollEmailsAsync(stoppingToken),
                    PollDiscordAsync(stoppingToken),
                    MonitorAlertsAsync(stoppingToken),
                    ComputeResponseTimeStatsAsync(stoppingToken),
                    CleanupAbandonedTicketsAsync(stoppingToken)
                };

                await Task.WhenAll(processes);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker shutting down gracefully");
            }
            catch (Exception ex)
            {
                _logger.LogCritical(ex, "Worker encountered a fatal error. Restarting loops in 30s...");
                await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);
            }
        }
    }

    private async Task PollEmailsAsync(CancellationToken stoppingToken)
    {
        // Brief startup delay to let MongoDB connections settle
        await Task.Delay(TimeSpan.FromSeconds(5), stoppingToken);

        var triggersCollection = _database.GetCollection<BsonDocument>(
            _appConfig.ZipStationMongoDb.Collections.WorkerTriggers);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                _logger.LogDebug("Email polling cycle starting");
                await _emailPollingService.PollAllProjectsAsync(stoppingToken);

                // Record poll completion
                await triggersCollection.InsertOneAsync(new BsonDocument
                {
                    { "type", "poll-complete" },
                    { "completedAt", DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() }
                }, cancellationToken: stoppingToken);

                // Clean up old poll-complete records (keep last 5)
                var oldRecords = await triggersCollection
                    .Find(new BsonDocument("type", "poll-complete"))
                    .SortByDescending(d => d["completedAt"])
                    .Skip(5)
                    .ToListAsync(stoppingToken);
                if (oldRecords.Count > 0)
                {
                    var oldIds = oldRecords.Select(r => r["_id"]).ToList();
                    await triggersCollection.DeleteManyAsync(
                        Builders<BsonDocument>.Filter.In("_id", oldIds), stoppingToken);
                }

                // Wait up to 2 minutes, but check for triggers every 5 seconds
                var waited = TimeSpan.Zero;
                var pollInterval = TimeSpan.FromMinutes(2);
                var checkInterval = TimeSpan.FromSeconds(5);

                while (waited < pollInterval && !stoppingToken.IsCancellationRequested)
                {
                    // Check for manual trigger
                    var trigger = await triggersCollection.FindOneAndDeleteAsync(
                        new BsonDocument("type", "poll-emails"),
                        cancellationToken: stoppingToken);

                    if (trigger != null)
                    {
                        _logger.LogInformation("Manual poll trigger received — polling now");
                        break;
                    }

                    // Check for import-history trigger
                    var historyTrigger = await triggersCollection.FindOneAndDeleteAsync(
                        new BsonDocument("type", "import-history"),
                        cancellationToken: stoppingToken);

                    if (historyTrigger != null)
                    {
                        _logger.LogInformation("History import trigger received — importing seen emails");
                        await _emailPollingService.ImportHistoryAsync(stoppingToken);
                    }

                    await Task.Delay(checkInterval, stoppingToken);
                    waited += checkInterval;
                }
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { break; }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in email polling loop");
                await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);
            }
        }
    }

    private async Task PollDiscordAsync(CancellationToken stoppingToken)
    {
        // Slight stagger from email poll so the two don't both bang Mongo at second 0.
        await Task.Delay(TimeSpan.FromSeconds(10), stoppingToken);

        var triggersCollection = _database.GetCollection<BsonDocument>(
            _appConfig.ZipStationMongoDb.Collections.WorkerTriggers);

        // Initial poll at startup so projects aren't waiting 60s for their first sync.
        try { await _discordPollingService.PollAllProjectsAsync(stoppingToken); }
        catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { return; }
        catch (Exception ex) { _logger.LogError(ex, "Initial Discord poll failed"); }

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                // Wait up to 60s for either the timeout or a manual sync-now trigger.
                // Exactly one poll runs per outer iteration: either the trigger's specific project,
                // or all projects on natural cadence. The previous loop polled at the top AND on
                // trigger which made every sync-now fire twice.
                string? triggeredProjectId = null;
                bool triggered = false;
                var waited = TimeSpan.Zero;
                var pollInterval = TimeSpan.FromSeconds(60);
                var checkInterval = TimeSpan.FromSeconds(5);

                while (waited < pollInterval && !stoppingToken.IsCancellationRequested)
                {
                    var trigger = await triggersCollection.FindOneAndDeleteAsync(
                        new BsonDocument("type", "poll-discord"),
                        cancellationToken: stoppingToken);
                    if (trigger != null)
                    {
                        triggeredProjectId = trigger.GetValue("projectId", BsonNull.Value)?.AsString;
                        triggered = true;
                        _logger.LogInformation("Discord sync trigger received{Scope}", triggeredProjectId is null ? "" : $" (project {triggeredProjectId})");
                        break;
                    }

                    await Task.Delay(checkInterval, stoppingToken);
                    waited += checkInterval;
                }

                if (triggered && !string.IsNullOrEmpty(triggeredProjectId))
                {
                    await _discordPollingService.PollProjectAsync(triggeredProjectId, stoppingToken);
                }
                else
                {
                    // Either the 60s tick fired, or the trigger was unscoped (poll everything).
                    _logger.LogDebug("Discord polling cycle starting");
                    await _discordPollingService.PollAllProjectsAsync(stoppingToken);
                }
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { break; }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in Discord polling loop");
                await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);
            }
        }
    }

    private async Task MonitorAlertsAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                _logger.LogDebug("Alert monitoring cycle starting");
                // TODO: Evaluate alert conditions, fire webhook notifications
                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { break; }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in alert monitoring loop");
                await Task.Delay(TimeSpan.FromSeconds(30), stoppingToken);
            }
        }
    }

    private async Task ComputeResponseTimeStatsAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                _logger.LogDebug("Response time stats computation starting");
                // TODO: Pre-compute per-agent and per-project response time analytics
                await Task.Delay(TimeSpan.FromHours(1), stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { break; }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in response time stats loop");
                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
            }
        }
    }

    private async Task CleanupAbandonedTicketsAsync(CancellationToken stoppingToken)
    {
        // Wait for initial startup
        await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                _logger.LogDebug("Abandoned ticket cleanup starting");

                var projects = await new Repositories.ProjectRepository(
                    _database, _appConfig.ZipStationMongoDb.Collections.Projects)
                    .GetAllWithImapAsync();

                // Also get projects without IMAP for the stale check
                var allProjects = await _database
                    .GetCollection<Entities.Project>(_appConfig.ZipStationMongoDb.Collections.Projects)
                    .Find(MongoDB.Driver.Builders<Entities.Project>.Filter.Eq(p => p.IsVoid, false))
                    .ToListAsync(stoppingToken);

                var ticketsCollection = _database.GetCollection<Entities.Ticket>(_appConfig.ZipStationMongoDb.Collections.Tickets);

                foreach (var project in allProjects)
                {
                    var staleThresholdDays = project.Settings?.StaleTicketDays > 0 ? project.Settings.StaleTicketDays : 5;
                    var cutoff = DateTimeOffset.UtcNow.AddDays(-staleThresholdDays).ToUnixTimeMilliseconds();

                    // Find open/pending tickets with no activity past the threshold, excluding preserved
                    var filter = MongoDB.Driver.Builders<Entities.Ticket>.Filter.Eq(t => t.ProjectId, project.Id)
                               & MongoDB.Driver.Builders<Entities.Ticket>.Filter.In(t => t.Status, new[] { 0, 1 }) // Open, Pending
                               & MongoDB.Driver.Builders<Entities.Ticket>.Filter.Eq(t => t.IsVoid, false)
                               & MongoDB.Driver.Builders<Entities.Ticket>.Filter.Ne(t => t.IsPreserved, true)
                               & MongoDB.Driver.Builders<Entities.Ticket>.Filter.Lt(t => t.UpdatedOnDateTime, cutoff);

                    var staleTickets = await ticketsCollection.Find(filter).ToListAsync(stoppingToken);

                    foreach (var ticket in staleTickets)
                    {
                        ticket.Status = 5; // Abandoned
                        ticket.UpdatedOnDateTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
                        await ticketsCollection.ReplaceOneAsync(
                            MongoDB.Driver.Builders<Entities.Ticket>.Filter.Eq(t => t.Id, ticket.Id),
                            ticket, cancellationToken: stoppingToken);

                        _logger.LogInformation("Ticket {TicketId} in project {ProjectName} marked as Abandoned (no activity for {Days} days)",
                            ticket.Id, project.Name, staleThresholdDays);
                    }
                }

                await Task.Delay(TimeSpan.FromHours(6), stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { break; }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in abandoned ticket cleanup loop");
                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
            }
        }
    }
}
