namespace ZipStation.Worker.Helpers;

public class AppConfig
{
    public string ApplicationId { get; set; } = "zipstation-worker";
    public string AppName { get; set; } = "ZipStationWorker";
    public string EncryptionKey { get; set; } = string.Empty;
    public ZipStationMongoDbConfiguration ZipStationMongoDb { get; set; } = new();
}

public class ZipStationMongoDbConfiguration
{
    public string ConnectionString { get; set; } = "mongodb://localhost:27017";
    public string DatabaseName { get; set; } = "zipstation";
    public ZipStationMongoDbCollections Collections { get; set; } = new();
}

public class ZipStationMongoDbCollections
{
    public string Projects { get; set; } = "projects";
    public string IntakeEmails { get; set; } = "intakeEmails";
    public string IntakeRules { get; set; } = "intakeRules";
    public string Tickets { get; set; } = "tickets";
    public string TicketMessages { get; set; } = "ticketMessages";
    public string Customers { get; set; } = "customers";
    public string TicketIdCounters { get; set; } = "ticketIdCounters";
    public string WorkerTriggers { get; set; } = "workerTriggers";
    public string Alerts { get; set; } = "alerts";
}
