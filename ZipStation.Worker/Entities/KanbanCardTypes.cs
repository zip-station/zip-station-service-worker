using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Bson.Serialization.Serializers;

namespace ZipStation.Worker.Entities;

/// Worker-side copy of the built-in story types. Kept in sync with the API's
/// <c>ZipStation.Models.Constants.KanbanCardTypes</c> (the worker has no project reference to
/// the API). A card's <c>Type</c> is a built-in name or a custom type's id.
public static class KanbanCardTypes
{
    public const string Feature = "Feature";
    public const string Bug = "Bug";
    public const string Improvement = "Improvement";
    public const string TechDebt = "TechDebt";

    /// Built-in names in historical enum order — index == legacy BSON int value.
    public static readonly IReadOnlyList<string> BuiltIns = new[] { Feature, Bug, Improvement, TechDebt };

    public static string? FromLegacyInt(int value) =>
        value >= 0 && value < BuiltIns.Count ? BuiltIns[value] : null;
}

/// Reads story-type fields that may still be stored as the legacy BSON int (0–3) and surfaces
/// them as strings, writing new values as plain strings. Mirrors the API serializer so the
/// worker can read legacy cards/sources without breaking and writes string types going forward.
public class LegacyCardTypeStringSerializer : SerializerBase<string?>
{
    public override string? Deserialize(BsonDeserializationContext context, BsonDeserializationArgs args)
    {
        var reader = context.Reader;
        switch (reader.GetCurrentBsonType())
        {
            case BsonType.String:
                return reader.ReadString();
            case BsonType.Int32:
                return KanbanCardTypes.FromLegacyInt(reader.ReadInt32()) ?? KanbanCardTypes.Feature;
            case BsonType.Int64:
                return KanbanCardTypes.FromLegacyInt((int)reader.ReadInt64()) ?? KanbanCardTypes.Feature;
            case BsonType.Null:
                reader.ReadNull();
                return null;
            default:
                reader.SkipValue();
                return null;
        }
    }

    public override void Serialize(BsonSerializationContext context, BsonSerializationArgs args, string? value)
    {
        if (value == null)
            context.Writer.WriteNull();
        else
            context.Writer.WriteString(value);
    }
}
