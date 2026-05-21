using MongoDB.Bson;
using MongoDB.Driver;
using ZipStation.Worker.Entities;

namespace ZipStation.Worker.Repositories;

public class ProjectRepository
{
    private readonly IMongoCollection<Project> _collection;
    public ProjectRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<Project>(collectionName);

    public async Task<List<Project>> GetAllWithImapAsync()
    {
        var filter = Builders<Project>.Filter.Eq(p => p.IsVoid, false)
                   & Builders<Project>.Filter.Ne(p => p.Settings.Imap, null);
        return await _collection.Find(filter).ToListAsync();
    }

    public async Task<Project?> GetByIdAsync(string id)
    {
        var filter = Builders<Project>.Filter.Eq(p => p.Id, id)
                   & Builders<Project>.Filter.Eq(p => p.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<List<Project>> GetAllWithDiscordEnabledAsync()
    {
        // Enabled flag implies the bot token has been set (enforced by API).
        var filter = Builders<Project>.Filter.Eq(p => p.IsVoid, false)
                   & Builders<Project>.Filter.Ne(p => p.Settings.Discord, null)
                   & Builders<Project>.Filter.Eq("settings.discord.enabled", true);
        return await _collection.Find(filter).ToListAsync();
    }

    /// Persist a Discord source's cursor after processing new threads/messages.
    /// Read-modify-write rather than a positional update: Mongo's NamedIdMemberConvention
    /// silently rewrites any nested `Id` property to BSON `_id`, which makes typed positional
    /// filters surprisingly hard to get right across both API and worker conventions.
    /// The worker is the only writer of `LastSeenId`, so a write-write race here can't happen
    /// — concurrent edits to other source fields (name/channel/etc.) come from the API and
    /// occur on a different field path, so a stale-read replace would clobber them. To guard
    /// against that, we re-read just before writing and only touch the cursor.
    public async Task UpdateDiscordSourceCursorAsync(string projectId, string sourceId, string lastSeenId)
    {
        var project = await GetByIdAsync(projectId);
        var source = project?.Settings?.Discord?.Sources?.FirstOrDefault(s => s.Id == sourceId);
        if (project is null || source is null) return;

        source.LastSeenId = lastSeenId;
        project.UpdatedOnDateTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var filter = Builders<Project>.Filter.Eq(p => p.Id, projectId);
        await _collection.ReplaceOneAsync(filter, project);
    }
}

public class IntakeEmailRepository
{
    private readonly IMongoCollection<IntakeEmail> _collection;
    public IntakeEmailRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<IntakeEmail>(collectionName);

    public async Task<IntakeEmail> CreateAsync(IntakeEmail entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    public async Task<IntakeEmail?> GetByMessageIdAsync(string messageId)
    {
        var filter = Builders<IntakeEmail>.Filter.Eq(e => e.MessageId, messageId)
                   & Builders<IntakeEmail>.Filter.Eq(e => e.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task UpdateAsync(IntakeEmail entity)
    {
        entity.UpdatedOnDateTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var filter = Builders<IntakeEmail>.Filter.Eq(e => e.Id, entity.Id);
        await _collection.ReplaceOneAsync(filter, entity);
    }
}

public class IntakeRuleRepository
{
    private readonly IMongoCollection<IntakeRule> _collection;
    public IntakeRuleRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<IntakeRule>(collectionName);

    public async Task<List<IntakeRule>> GetEnabledByProjectIdAsync(string projectId)
    {
        var filter = Builders<IntakeRule>.Filter.Eq(r => r.ProjectId, projectId)
                   & Builders<IntakeRule>.Filter.Eq(r => r.IsEnabled, true)
                   & Builders<IntakeRule>.Filter.Eq(r => r.IsVoid, false);
        return await _collection.Find(filter).SortBy(r => r.Priority).ToListAsync();
    }
}

public class TicketRepository
{
    private readonly IMongoCollection<Ticket> _collection;
    public TicketRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<Ticket>(collectionName);

    public async Task<Ticket> CreateAsync(Ticket entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        if (entity.CreatedOnDateTime == 0)
            entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    public async Task<Ticket?> GetByIdAsync(string id)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.Id, id)
                   & Builders<Ticket>.Filter.Eq(t => t.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<Ticket?> GetByCustomerEmailAndProjectAsync(string email, string projectId)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.CustomerEmail, email)
                   & Builders<Ticket>.Filter.Eq(t => t.ProjectId, projectId)
                   & Builders<Ticket>.Filter.Eq(t => t.IsVoid, false)
                   & Builders<Ticket>.Filter.In(t => t.Status, new[] { 0, 1 }); // Open, Pending
        return await _collection.Find(filter).SortByDescending(t => t.CreatedOnDateTime).FirstOrDefaultAsync();
    }

    public async Task<Ticket?> GetByTicketNumberAndProjectAsync(long ticketNumber, string projectId)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.TicketNumber, ticketNumber)
                   & Builders<Ticket>.Filter.Eq(t => t.ProjectId, projectId)
                   & Builders<Ticket>.Filter.Eq(t => t.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task UpdateStatusAsync(string ticketId, int status)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.Id, ticketId);
        var update = Builders<Ticket>.Update
            .Set(t => t.Status, status)
            .Set(t => t.UpdatedOnDateTime, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
        await _collection.UpdateOneAsync(filter, update);
    }

    public async Task SetLastMessageSourceAsync(string ticketId, int source)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.Id, ticketId);
        var update = Builders<Ticket>.Update
            .Set(t => t.LastMessageSource, source)
            .Set(t => t.UpdatedOnDateTime, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
        await _collection.UpdateOneAsync(filter, update);
    }

    public async Task<bool> ExistsByTicketNumberAndProjectAsync(string projectId, long ticketNumber)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.ProjectId, projectId)
                   & Builders<Ticket>.Filter.Eq(t => t.TicketNumber, ticketNumber);
        return await _collection.CountDocumentsAsync(filter) > 0;
    }

    public async Task<List<Ticket>> GetRecentOpenByProjectIdAsync(string projectId, int limit, string excludeTicketId)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.ProjectId, projectId)
                   & Builders<Ticket>.Filter.Eq(t => t.IsVoid, false)
                   & Builders<Ticket>.Filter.In(t => t.Status, new[] { 0, 1 })
                   & Builders<Ticket>.Filter.Ne(t => t.Id, excludeTicketId);
        return await _collection.Find(filter)
            .SortByDescending(t => t.UpdatedOnDateTime)
            .Limit(limit)
            .ToListAsync();
    }
}

public class TicketMessageRepository
{
    private readonly IMongoCollection<TicketMessage> _collection;
    public TicketMessageRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<TicketMessage>(collectionName);

    public async Task<TicketMessage> CreateAsync(TicketMessage entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        if (entity.CreatedOnDateTime == 0)
            entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    public async Task UpdateAsync(TicketMessage entity)
    {
        entity.UpdatedOnDateTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        await _collection.ReplaceOneAsync(
            Builders<TicketMessage>.Filter.Eq(e => e.Id, entity.Id), entity);
    }

    public async Task<List<TicketMessage>> GetByTicketIdAsync(string ticketId)
    {
        var filter = Builders<TicketMessage>.Filter.Eq(m => m.TicketId, ticketId)
                   & Builders<TicketMessage>.Filter.Eq(m => m.IsVoid, false);
        return await _collection.Find(filter)
            .SortBy(m => m.CreatedOnDateTime)
            .ToListAsync();
    }
}

public class CustomerRepository
{
    private readonly IMongoCollection<Customer> _collection;
    public CustomerRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<Customer>(collectionName);

    public async Task<Customer?> GetByEmailAndProjectAsync(string email, string projectId)
    {
        var filter = Builders<Customer>.Filter.Eq(c => c.Email, email)
                   & Builders<Customer>.Filter.Eq(c => c.ProjectId, projectId)
                   & Builders<Customer>.Filter.Eq(c => c.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<Customer> CreateAsync(Customer entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    public async Task UpdateAsync(Customer entity)
    {
        entity.UpdatedOnDateTime = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var filter = Builders<Customer>.Filter.Eq(c => c.Id, entity.Id);
        await _collection.ReplaceOneAsync(filter, entity);
    }
}

public class TicketIdCounterRepository
{
    private readonly IMongoCollection<TicketIdCounter> _collection;
    public TicketIdCounterRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<TicketIdCounter>(collectionName);

    public async Task<long> GetNextTicketNumberAsync(string projectId)
    {
        var filter = Builders<TicketIdCounter>.Filter.Eq(c => c.ProjectId, projectId);
        var update = Builders<TicketIdCounter>.Update.Inc(c => c.CurrentValue, 1);
        var options = new FindOneAndUpdateOptions<TicketIdCounter>
        {
            IsUpsert = true,
            ReturnDocument = ReturnDocument.After
        };
        var result = await _collection.FindOneAndUpdateAsync(filter, update, options);
        return result.CurrentValue;
    }
}

public class MaxInstructionRepository
{
    private readonly IMongoCollection<MaxInstruction> _collection;
    public MaxInstructionRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<MaxInstruction>(collectionName);

    public async Task<List<MaxInstruction>> GetByProjectIdAsync(string projectId)
    {
        var filter = Builders<MaxInstruction>.Filter.Eq(i => i.ProjectId, projectId)
                   & Builders<MaxInstruction>.Filter.Eq(i => i.IsVoid, false);
        return await _collection.Find(filter).ToListAsync();
    }
}

public class MaxExampleReplyRepository
{
    private readonly IMongoCollection<MaxExampleReply> _collection;
    public MaxExampleReplyRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<MaxExampleReply>(collectionName);

    public async Task<List<MaxExampleReply>> GetByProjectIdAsync(string projectId)
    {
        var filter = Builders<MaxExampleReply>.Filter.Eq(r => r.ProjectId, projectId)
                   & Builders<MaxExampleReply>.Filter.Eq(r => r.IsVoid, false);
        return await _collection.Find(filter).ToListAsync();
    }
}

public class MaxTicketEnrichmentRepository
{
    private readonly IMongoCollection<MaxTicketEnrichment> _collection;
    public MaxTicketEnrichmentRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<MaxTicketEnrichment>(collectionName);

    public async Task<MaxTicketEnrichment?> GetByTicketIdAsync(string ticketId)
    {
        var filter = Builders<MaxTicketEnrichment>.Filter.Eq(e => e.TicketId, ticketId)
                   & Builders<MaxTicketEnrichment>.Filter.Eq(e => e.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<List<MaxTicketEnrichment>> GetRecentByProjectIdAsync(string projectId, int limit)
    {
        var filter = Builders<MaxTicketEnrichment>.Filter.Eq(e => e.ProjectId, projectId)
                   & Builders<MaxTicketEnrichment>.Filter.Eq(e => e.IsVoid, false);
        return await _collection.Find(filter)
            .SortByDescending(e => e.CreatedOnDateTime)
            .Limit(limit)
            .ToListAsync();
    }

    public async Task<MaxTicketEnrichment> UpsertAsync(MaxTicketEnrichment entity)
    {
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        entity.UpdatedOnDateTime = now;

        // Preserve existing _id (Mongo treats it as immutable on replace) and CreatedOnDateTime.
        // BaseEntity ctor auto-generates a fresh ObjectId, so callers passing a brand-new entity
        // would otherwise hit error code 66 when a doc for this TicketId already exists.
        var filter = Builders<MaxTicketEnrichment>.Filter.Eq(e => e.TicketId, entity.TicketId);
        var existing = await _collection.Find(filter).FirstOrDefaultAsync();
        if (existing != null)
        {
            entity.Id = existing.Id;
            entity.CreatedOnDateTime = existing.CreatedOnDateTime;
        }
        else
        {
            if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
            if (entity.CreatedOnDateTime == 0) entity.CreatedOnDateTime = now;
        }

        var options = new ReplaceOptions { IsUpsert = true };
        await _collection.ReplaceOneAsync(filter, entity, options);
        return entity;
    }
}

public class MaxTaskRepository
{
    private readonly IMongoCollection<MaxTask> _collection;
    public MaxTaskRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<MaxTask>(collectionName);

    public async Task<MaxTask> CreateAsync(MaxTask entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    public async Task<long> SoftDeletePendingByTicketIdAsync(string ticketId)
    {
        var filter = Builders<MaxTask>.Filter.Eq(t => t.TicketId, ticketId)
                   & Builders<MaxTask>.Filter.Eq(t => t.Status, "pending")
                   & Builders<MaxTask>.Filter.Eq(t => t.IsVoid, false);
        var update = Builders<MaxTask>.Update
            .Set(t => t.IsVoid, true)
            .Set(t => t.UpdatedOnDateTime, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
        var result = await _collection.UpdateManyAsync(filter, update);
        return result.ModifiedCount;
    }
}

public class MaxStoryEnrichmentRepository
{
    private readonly IMongoCollection<MaxStoryEnrichment> _collection;
    public MaxStoryEnrichmentRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<MaxStoryEnrichment>(collectionName);

    public async Task<MaxStoryEnrichment?> GetByStoryIdAsync(string storyId)
    {
        var filter = Builders<MaxStoryEnrichment>.Filter.Eq(e => e.StoryId, storyId)
                   & Builders<MaxStoryEnrichment>.Filter.Eq(e => e.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<MaxStoryEnrichment> UpsertAsync(MaxStoryEnrichment entity)
    {
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        entity.UpdatedOnDateTime = now;

        // Preserve the existing _id when a doc for this StoryId is already in Mongo.
        // BaseEntity ctor auto-generates a fresh ObjectId, and Mongo treats _id as immutable
        // on replace, so a naive ReplaceOne would fail with code 66.
        var filter = Builders<MaxStoryEnrichment>.Filter.Eq(e => e.StoryId, entity.StoryId);
        var existing = await _collection.Find(filter).FirstOrDefaultAsync();
        if (existing != null)
        {
            entity.Id = existing.Id;
            entity.CreatedOnDateTime = existing.CreatedOnDateTime;
        }
        else
        {
            if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
            if (entity.CreatedOnDateTime == 0) entity.CreatedOnDateTime = now;
        }

        var options = new ReplaceOptions { IsUpsert = true };
        await _collection.ReplaceOneAsync(filter, entity, options);
        return entity;
    }
}

public class KanbanBoardRepository
{
    private readonly IMongoCollection<KanbanBoard> _collection;
    public KanbanBoardRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<KanbanBoard>(collectionName);

    public async Task<KanbanBoard?> GetByProjectIdAsync(string projectId)
    {
        var filter = Builders<KanbanBoard>.Filter.Eq(b => b.ProjectId, projectId)
                   & Builders<KanbanBoard>.Filter.Eq(b => b.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }
}

public class KanbanCardRepository
{
    private readonly IMongoCollection<KanbanCard> _collection;
    public KanbanCardRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<KanbanCard>(collectionName);

    public async Task<KanbanCard> CreateAsync(KanbanCard entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        if (entity.CreatedOnDateTime == 0)
            entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    /// Idempotency: skip a Discord thread/message we've already converted to a card.
    /// Matches a single ExternalSources entry on both type and messageId via $elemMatch.
    public async Task<KanbanCard?> GetByExternalMessageIdAsync(string projectId, int sourceType, string messageId)
    {
        var elemMatch = Builders<KanbanCardExternalSource>.Filter.Eq(s => s.Type, sourceType)
                      & Builders<KanbanCardExternalSource>.Filter.Eq(s => s.MessageId, messageId);
        var filter = Builders<KanbanCard>.Filter.Eq(c => c.ProjectId, projectId)
                   & Builders<KanbanCard>.Filter.Eq(c => c.IsVoid, false)
                   & Builders<KanbanCard>.Filter.ElemMatch(c => c.ExternalSources, elemMatch);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<double> GetMaxPositionInColumnAsync(string boardId, string columnId)
    {
        var filter = Builders<KanbanCard>.Filter.Eq(c => c.BoardId, boardId)
                   & Builders<KanbanCard>.Filter.Eq(c => c.ColumnId, columnId)
                   & Builders<KanbanCard>.Filter.Eq(c => c.IsVoid, false);
        var top = await _collection.Find(filter)
            .SortByDescending(c => c.Position)
            .Limit(1)
            .FirstOrDefaultAsync();
        return top?.Position ?? 0;
    }

    public async Task<KanbanCard?> GetByCardNumberAsync(string projectId, long cardNumber)
    {
        var filter = Builders<KanbanCard>.Filter.Eq(c => c.ProjectId, projectId)
                   & Builders<KanbanCard>.Filter.Eq(c => c.CardNumber, cardNumber)
                   & Builders<KanbanCard>.Filter.Eq(c => c.IsVoid, false);
        return await _collection.Find(filter).FirstOrDefaultAsync();
    }

    public async Task<List<KanbanCard>> GetByTicketIdAsync(string ticketId)
    {
        var filter = Builders<KanbanCard>.Filter.AnyEq(c => c.LinkedTicketIds, ticketId)
                   & Builders<KanbanCard>.Filter.Eq(c => c.IsVoid, false);
        return await _collection.Find(filter).ToListAsync();
    }

    public async Task<List<KanbanCard>> GetByAnyLinkedTicketIdAsync(IEnumerable<string> ticketIds)
    {
        var ids = ticketIds?.Where(id => !string.IsNullOrEmpty(id)).Distinct().ToList() ?? new List<string>();
        if (ids.Count == 0) return new List<KanbanCard>();
        var filter = Builders<KanbanCard>.Filter.AnyIn(c => c.LinkedTicketIds, ids)
                   & Builders<KanbanCard>.Filter.Eq(c => c.IsVoid, false);
        return await _collection.Find(filter).ToListAsync();
    }
}

public class KanbanCardNumberCounterRepository
{
    private readonly IMongoCollection<KanbanCardNumberCounter> _collection;
    public KanbanCardNumberCounterRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<KanbanCardNumberCounter>(collectionName);

    public async Task<long> GetNextCardNumberAsync(string projectId)
    {
        var filter = Builders<KanbanCardNumberCounter>.Filter.Eq(c => c.ProjectId, projectId);
        var update = Builders<KanbanCardNumberCounter>.Update.Inc(c => c.CurrentValue, 1);
        var options = new FindOneAndUpdateOptions<KanbanCardNumberCounter>
        {
            IsUpsert = true,
            ReturnDocument = ReturnDocument.After
        };
        var result = await _collection.FindOneAndUpdateAsync(filter, update, options);
        return result.CurrentValue;
    }
}

public class MaxQuestionRepository
{
    private readonly IMongoCollection<MaxQuestion> _collection;
    public MaxQuestionRepository(IMongoDatabase db, string collectionName) => _collection = db.GetCollection<MaxQuestion>(collectionName);

    public async Task<MaxQuestion> CreateAsync(MaxQuestion entity)
    {
        if (string.IsNullOrEmpty(entity.Id)) entity.Id = ObjectId.GenerateNewId().ToString();
        var now = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
    }

    public async Task<long> SoftDeletePendingByTicketIdAsync(string ticketId)
    {
        var filter = Builders<MaxQuestion>.Filter.Eq(q => q.SourceTicketId, ticketId)
                   & Builders<MaxQuestion>.Filter.Eq(q => q.Status, "pending")
                   & Builders<MaxQuestion>.Filter.Eq(q => q.IsVoid, false);
        var update = Builders<MaxQuestion>.Update
            .Set(q => q.IsVoid, true)
            .Set(q => q.UpdatedOnDateTime, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds());
        var result = await _collection.UpdateManyAsync(filter, update);
        return result.ModifiedCount;
    }
}
