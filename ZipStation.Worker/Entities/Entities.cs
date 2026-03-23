using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace ZipStation.Worker.Entities;

[BsonIgnoreExtraElements]
public class BaseEntity
{
    [BsonId]
    [BsonElement("_id")]
    [BsonRepresentation(BsonType.String)]
    public string Id { get; set; } = ObjectId.GenerateNewId().ToString();
    public long UpdatedOnDateTime { get; set; }
    public long CreatedOnDateTime { get; set; }
    public bool IsVoid { get; set; }
}

[BsonIgnoreExtraElements]
public class Project : BaseEntity
{
    public string CompanyId { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Slug { get; set; } = string.Empty;
    public string SupportEmailAddress { get; set; } = string.Empty;
    public ProjectSettings Settings { get; set; } = new();
}

[BsonIgnoreExtraElements]
public class ProjectSettings
{
    public ImapSettings? Imap { get; set; }
    public SmtpSettings? Smtp { get; set; }
    public TicketIdSettings TicketId { get; set; } = new();
    public int StaleTicketDays { get; set; } = 5;
    public SpamSettings? Spam { get; set; }
    public ContactFormSettings? ContactForm { get; set; }
}

[BsonIgnoreExtraElements]
public class SpamSettings
{
    public int AutoDenyThreshold { get; set; } = 80;
    public int FlagThreshold { get; set; } = 50;
    public bool AutoDenyEnabled { get; set; }
}

[BsonIgnoreExtraElements]
public class ContactFormSettings
{
    public bool Enabled { get; set; }
    public List<string> SystemSenderEmails { get; set; } = new();
    public string EmailLabel { get; set; } = "Email";
    public string NameLabel { get; set; } = "Name";
    public string MessageLabel { get; set; } = "Message";
    public string? SubjectLabel { get; set; }
}

[BsonIgnoreExtraElements]
public class ImapSettings
{
    public string Host { get; set; } = string.Empty;
    public int Port { get; set; } = 993;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public bool UseSsl { get; set; } = true;
}

[BsonIgnoreExtraElements]
public class SmtpSettings
{
    public string Host { get; set; } = string.Empty;
    public int Port { get; set; } = 587;
    public string Username { get; set; } = string.Empty;
    public string Password { get; set; } = string.Empty;
    public bool UseSsl { get; set; } = true;
    public string? FromName { get; set; }
    public string? FromEmail { get; set; }
}

[BsonIgnoreExtraElements]
public class TicketIdSettings
{
    public string Prefix { get; set; } = string.Empty;
    public int MinLength { get; set; } = 3;
    public int MaxLength { get; set; } = 6;
    public int Format { get; set; } = 0; // 0=Numeric, 1=Alphanumeric, 2=DateNumeric
    public string SubjectTemplate { get; set; } = "{ProjectName} - Ticket {TicketId}";
}

public class IntakeEmail : BaseEntity
{
    public string CompanyId { get; set; } = string.Empty;
    public string ProjectId { get; set; } = string.Empty;
    public string FromEmail { get; set; } = string.Empty;
    public string FromName { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string BodyText { get; set; } = string.Empty;
    public string? BodyHtml { get; set; }
    public long ReceivedOn { get; set; }
    public int Status { get; set; } = 0; // 0=Pending, 1=Approved, 2=Denied
    public int SpamScore { get; set; }
    public bool DeniedPermanently { get; set; }
    public string? ApprovedByUserId { get; set; }
    public string? DeniedByUserId { get; set; }
    public long ProcessedOn { get; set; }
    public string? TicketId { get; set; }
    public string? MessageId { get; set; }
    public string? InReplyTo { get; set; }
}

[BsonIgnoreExtraElements]
public class IntakeRule : BaseEntity
{
    public string CompanyId { get; set; } = string.Empty;
    public string ProjectId { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public List<IntakeRuleCondition> Conditions { get; set; } = new();
    public int Action { get; set; } // 0=AutoApprove, 1=AutoDeny, 2=AutoDenyPermanent
    public int Priority { get; set; }
    public bool IsEnabled { get; set; } = true;
}

[BsonIgnoreExtraElements]
public class IntakeRuleCondition
{
    public int Type { get; set; } // 0=FromEmail, 1=FromDomain, 2=SubjectContains, 3=BodyContains
    public string Value { get; set; } = string.Empty;
}

[BsonIgnoreExtraElements]
public class Ticket : BaseEntity
{
    public string CompanyId { get; set; } = string.Empty;
    public string ProjectId { get; set; } = string.Empty;
    public long TicketNumber { get; set; }
    public string Subject { get; set; } = string.Empty;
    public int Status { get; set; } = 0; // 0=Open, 1=Pending, 2=Resolved, 3=Closed, 4=Merged
    public int Priority { get; set; } = 1; // 0=Low, 1=Normal, 2=High, 3=Urgent
    public string? AssignedToUserId { get; set; }
    public string? CustomerName { get; set; }
    public string? CustomerEmail { get; set; }
    public List<string> Tags { get; set; } = new();
    public int CreationSource { get; set; } = 0;
}

public class TicketMessage : BaseEntity
{
    public string TicketId { get; set; } = string.Empty;
    public string CompanyId { get; set; } = string.Empty;
    public string ProjectId { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public string? BodyHtml { get; set; }
    public bool IsInternalNote { get; set; }
    public string? AuthorUserId { get; set; }
    public string? AuthorName { get; set; }
    public string? AuthorEmail { get; set; }
    public int Source { get; set; } = 1; // 0=Customer, 1=Agent, 2=System
}

[BsonIgnoreExtraElements]
public class Customer : BaseEntity
{
    public string CompanyId { get; set; } = string.Empty;
    public string ProjectId { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public List<string> Tags { get; set; } = new();
    public string? Notes { get; set; }
    public bool IsBanned { get; set; }
    public int OpenTicketCount { get; set; }
    public int ClosedTicketCount { get; set; }
    public int TotalTicketCount { get; set; }
}

[BsonIgnoreExtraElements]
public class Alert
{
    [BsonId]
    [BsonElement("_id")]
    [BsonRepresentation(BsonType.String)]
    public string Id { get; set; } = ObjectId.GenerateNewId().ToString();
    public string ProjectId { get; set; } = string.Empty;
    public int TriggerType { get; set; } // 0=NewTicket
    public string? TriggerValue { get; set; }
    public int ChannelType { get; set; } // 0=Slack, 1=Discord, 2=Generic
    public string WebhookUrl { get; set; } = string.Empty;
    public bool IsEnabled { get; set; } = true;
    public bool IsVoid { get; set; }
}

public class TicketIdCounter
{
    [BsonId]
    [BsonElement("_id")]
    public string ProjectId { get; set; } = string.Empty;
    public long CurrentValue { get; set; }
}
