using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OutlookArchive.UI;

internal static class RemovalPreferences
{
    private static string FileFor(string root, string job, string? directory)
    {
        string identity = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar).ToUpperInvariant() + "|" + job;
        string key = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
        return Path.Combine(directory ?? Path.Combine(root,"preferences","removal-options"), key + ".json");
    }
    internal static bool? Read(string root, string job, string? directory = null)
    {
        try { string file = FileFor(root, job, directory); return File.Exists(file) ? JsonSerializer.Deserialize<bool>(File.ReadAllText(file)) : null; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return null; }
    }
    internal static void Save(string root, string job, bool allowed, string? directory = null)
    {
        string file = FileFor(root, job, directory); Directory.CreateDirectory(Path.GetDirectoryName(file)!);
        string temporary = file + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try { File.WriteAllText(temporary, JsonSerializer.Serialize(allowed)); File.Move(temporary, file, true); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
