using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace EveX.Settings.Bridge;

public static class InstallationIdentity
{
    public static string Normalize(string directory)
    {
        var path = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        // A drive root must remain absolute when passed to Path.GetFullPath again.
        if (path.Length == 2 && path[1] == ':') path += Path.DirectorySeparatorChar;
        return path.ToLowerInvariant();
    }
    public static string Hash(string directory) => Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(Normalize(directory).TrimEnd('\\', '/'))));
}

public sealed class BridgeException(string code, string message, JsonNode? details = null, bool fromServer = false) : Exception(message)
{
    public string Code { get; } = code;
    public JsonNode? Details { get; } = details;
    public bool FromServer { get; } = fromServer;
    public bool IsDisconnected => Code is "disconnected" or "timeout" or "not-ready" or "stale-session" or "runtime-started" or "uncertain-apply";
}

public static class Protocol
{
    public const int Version = 1;
    public const int MaximumBytes = 4 * 1024 * 1024;
    public const nuint CopyDataIdentifier = 0x45585031;
    public static JsonObject Request(string method, string session, JsonObject? arguments = null, string? id = null) => new()
    {
        ["version"] = Version, ["id"] = id ?? Guid.NewGuid().ToString("N"), ["method"] = method,
        ["session"] = session, ["arguments"] = arguments ?? new JsonObject()
    };
    public static JsonObject Response(string text, string expectedId)
    {
        if (Encoding.UTF8.GetByteCount(text) > MaximumBytes) throw new BridgeException("invalid-response", "The settings response exceeds the size limit.");
        JsonObject response;
        try { response = JsonNode.Parse(text)?.AsObject() ?? throw new JsonException(); }
        catch (Exception e) when (e is JsonException or InvalidOperationException) { throw new BridgeException("invalid-response", "The settings service returned an incomplete response."); }
        if (response.Int("version") != Version || response.Text("id") != expectedId)
            throw new BridgeException("incompatible", "The settings response does not match this request or bridge version.");
        if (!response.Bool("ok"))
        {
            var error = response["error"];
            throw new BridgeException(error.Text("code", "service-error"), error.Text("message", "The settings operation failed."), error?["details"]?.DeepClone(), fromServer: true);
        }
        return response["data"] as JsonObject ?? new JsonObject { ["result"] = response["data"]?.DeepClone() };
    }
    public static void VerifyHello(JsonObject hello)
    {
        if (hello.Int("bridgeVersion") != Version || hello.Int("schemaVersion") != Version)
            throw new BridgeException("incompatible", "This settings editor and the application use incompatible settings versions. Use the internal editor or install a matching package.");
        if (!hello.Bool("ready")) throw new BridgeException("not-ready", "EVE-X-Preview is still loading its settings.");
    }
}

public static class JsonExtensions
{
    public static string Text(this JsonNode? node, string key, string fallback = "") => node?[key] is { } value ? Display(value) : fallback;
    public static bool Bool(this JsonNode? node, string key, bool fallback = false) => node?[key] is { } value ? IsTrue(value) : fallback;
    public static int Int(this JsonNode? node, string key, int fallback = 0) => int.TryParse(node.Text(key), out var value) ? value : fallback;
    public static double Number(this JsonNode? node, string key, double fallback = 0) => double.TryParse(node.Text(key), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var value) ? value : fallback;
    public static JsonArray Array(this JsonNode? node, string key) => node?[key] as JsonArray ?? new JsonArray();
    public static JsonObject Object(this JsonNode? node, string key) => node?[key] as JsonObject ?? new JsonObject();
    public static string Display(JsonNode? node) => node is JsonValue value && value.TryGetValue<string>(out var text) ? text : node?.ToJsonString() ?? "";
    public static bool IsTrue(JsonNode? node) => Display(node) is "true" or "1";
    public static JsonObject Copy(this JsonObject node) => (JsonObject)node.DeepClone();
}

public sealed class Metadata(JsonObject value)
{
    public JsonObject Value { get; } = value;
    public IEnumerable<JsonObject> Sections => Value.Array("sections").OfType<JsonObject>();
    public IEnumerable<JsonObject> Fields => Value["fields"] is JsonObject keyed ? keyed.Select(x => x.Value).OfType<JsonObject>() : Value.Array("fields").OfType<JsonObject>();
    public IEnumerable<JsonObject> Commands => Value.Array("commands").OfType<JsonObject>();
    public IEnumerable<JsonObject> Events => Value.Array("events").OfType<JsonObject>();
    public IEnumerable<JsonObject> Tables => Value.Array("tables").OfType<JsonObject>();
    public JsonObject? Field(string id) => Fields.FirstOrDefault(f => f.Text("id") == id);
    public IEnumerable<JsonObject> SectionFields(JsonObject section) => section.Array("fields").Select(id => Field(JsonExtensions.Display(id))).OfType<JsonObject>();
}
