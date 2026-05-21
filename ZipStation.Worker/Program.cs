using MongoDB.Bson.Serialization.Conventions;
using MongoDB.Driver;
using Serilog;
using Serilog.Events;
using ZipStation.Worker;
using ZipStation.Worker.Helpers;
using ZipStation.Worker.Repositories;
using ZipStation.Worker.Services;

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
    .MinimumLevel.Information()
    .Enrich.FromLogContext()
    .Enrich.WithProperty("Application", "ZipStationWorker")
    .WriteTo.Console(outputTemplate: "[{Timestamp:HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}")
    .CreateLogger();

try
{
    var builder = Host.CreateApplicationBuilder(args);
    builder.Logging.ClearProviders();
    builder.Services.AddSerilog();

    // Config
    var appConfig = builder.Configuration.Get<AppConfig>() ?? new AppConfig();
    builder.Services.Configure<AppConfig>(builder.Configuration);

    // Encryption
    if (!string.IsNullOrEmpty(appConfig.EncryptionKey))
        EncryptionHelper.Initialize(appConfig.EncryptionKey);

    // MongoDB
    var camelCaseConventionPack = new ConventionPack { new CamelCaseElementNameConvention() };
    ConventionRegistry.Register("camelCase", camelCaseConventionPack, type => true);

    var mongoSettings = MongoClientSettings.FromConnectionString(appConfig.ZipStationMongoDb.ConnectionString);
    mongoSettings.MaxConnectionPoolSize = 50;
    mongoSettings.MinConnectionPoolSize = 5;
    var client = new MongoClient(mongoSettings);
    var database = client.GetDatabase(appConfig.ZipStationMongoDb.DatabaseName);
    builder.Services.AddSingleton<IMongoClient>(client);
    builder.Services.AddSingleton(database);

    // Repositories
    var collections = appConfig.ZipStationMongoDb.Collections;
    builder.Services.AddSingleton(sp => new ProjectRepository(sp.GetRequiredService<IMongoDatabase>(), collections.Projects));
    builder.Services.AddSingleton(sp => new IntakeEmailRepository(sp.GetRequiredService<IMongoDatabase>(), collections.IntakeEmails));
    builder.Services.AddSingleton(sp => new IntakeRuleRepository(sp.GetRequiredService<IMongoDatabase>(), collections.IntakeRules));
    builder.Services.AddSingleton(sp => new TicketRepository(sp.GetRequiredService<IMongoDatabase>(), collections.Tickets));
    builder.Services.AddSingleton(sp => new TicketMessageRepository(sp.GetRequiredService<IMongoDatabase>(), collections.TicketMessages));
    builder.Services.AddSingleton(sp => new CustomerRepository(sp.GetRequiredService<IMongoDatabase>(), collections.Customers));
    builder.Services.AddSingleton(sp => new TicketIdCounterRepository(sp.GetRequiredService<IMongoDatabase>(), collections.TicketIdCounters));
    builder.Services.AddSingleton(sp => new MaxInstructionRepository(sp.GetRequiredService<IMongoDatabase>(), collections.MaxInstructions));
    builder.Services.AddSingleton(sp => new MaxExampleReplyRepository(sp.GetRequiredService<IMongoDatabase>(), collections.MaxExampleReplies));
    builder.Services.AddSingleton(sp => new MaxTicketEnrichmentRepository(sp.GetRequiredService<IMongoDatabase>(), collections.MaxTicketEnrichments));
    builder.Services.AddSingleton(sp => new MaxTaskRepository(sp.GetRequiredService<IMongoDatabase>(), collections.MaxTasks));
    builder.Services.AddSingleton(sp => new MaxQuestionRepository(sp.GetRequiredService<IMongoDatabase>(), collections.MaxQuestions));
    builder.Services.AddSingleton(sp => new KanbanBoardRepository(sp.GetRequiredService<IMongoDatabase>(), collections.KanbanBoards));
    builder.Services.AddSingleton(sp => new KanbanCardRepository(sp.GetRequiredService<IMongoDatabase>(), collections.KanbanCards));
    builder.Services.AddSingleton(sp => new KanbanCardNumberCounterRepository(sp.GetRequiredService<IMongoDatabase>(), collections.KanbanCardNumberCounters));
    builder.Services.AddSingleton(sp => new MaxStoryEnrichmentRepository(sp.GetRequiredService<IMongoDatabase>(), collections.MaxStoryEnrichments));

    // Services
    builder.Services.AddSingleton<FileStorageService>();
    builder.Services.AddSingleton<IEmailPollingService, EmailPollingService>();
    builder.Services.AddSingleton<IMaxEnrichmentService, MaxEnrichmentService>();
    builder.Services.AddSingleton<IMaxStoryTriageService, MaxStoryTriageService>();
    builder.Services.AddSingleton<IDiscordPollingService, DiscordPollingService>();

    // Worker
    builder.Services.AddHostedService<Worker>();

    var host = builder.Build();
    host.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Worker terminated unexpectedly");
}
finally
{
    Log.CloseAndFlush();
}
