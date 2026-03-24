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

    public async Task<bool> ExistsByTicketNumberAndProjectAsync(string projectId, long ticketNumber)
    {
        var filter = Builders<Ticket>.Filter.Eq(t => t.ProjectId, projectId)
                   & Builders<Ticket>.Filter.Eq(t => t.TicketNumber, ticketNumber);
        return await _collection.CountDocumentsAsync(filter) > 0;
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
        entity.CreatedOnDateTime = now;
        entity.UpdatedOnDateTime = now;
        entity.IsVoid = false;
        await _collection.InsertOneAsync(entity);
        return entity;
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
