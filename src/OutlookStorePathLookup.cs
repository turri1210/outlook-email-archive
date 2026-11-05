using System.Diagnostics;
using System.Text;
using System.Text.Json;

internal sealed partial class OutlookArchiveBackend
{
    private string? ResolveStoreIdFromOutlookMetadata(string path)
    {
        // Newly attached PSTs can be absent from a separate MAPI session's
        // table. Read only wrapped identifiers from Outlook's live collection;
        // never ask an unrelated store for FilePath or its root folder.
        object? stores = null; var matches = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            stores = ((dynamic)session!).Stores;
            for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
            {
                object? store = null;
                try
                {
                    store = ((dynamic)stores)[i]; string id = (string)((dynamic)store).StoreID;
                    if (FolderStructureRules.StoreIdContainsPath(id, path)) matches.Add(id);
                }
                catch (System.Runtime.InteropServices.COMException) { }
                finally { Release(store); }
            }
        }
        finally { Release(stores); }
        if (matches.Count > 1) throw new ArchiveReviewException("Mais de um identificador corresponde ao PST solicitado.");
        return matches.SingleOrDefault();
    }
    private static string? ResolveStoreIdByPath(string path)
    {
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath)) return null;
        string helper = Path.Combine(AppContext.BaseDirectory, "native-identity.exe");
        if (!File.Exists(helper)) throw new ArchiveReviewException("Componente de localização do PST ausente.");
        string request = Path.Combine(Path.GetTempPath(), "outlook-store-path-" + Guid.NewGuid().ToString("N") + ".txt");
        try
        {
            File.WriteAllText(request, "readonly-store-path-v1\n" + fullPath + "\n", new UTF8Encoding(false));
            using var process = new Process { StartInfo = new ProcessStartInfo(helper)
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8 } };
            process.StartInfo.ArgumentList.Add(request);
            process.Start(); var output = process.StandardOutput.ReadToEndAsync(); var errors = process.StandardError.ReadToEndAsync(); process.WaitForExit();
            string json = output.GetAwaiter().GetResult(), error = errors.GetAwaiter().GetResult();
            if (process.ExitCode != 0) throw new ArchiveReviewException("Não foi possível localizar apenas o PST solicitado: " + error.Trim());
            using var document = JsonDocument.Parse(json); var result = document.RootElement;
            if (result.GetProperty("Stage").GetString() != "complete" || result.GetProperty("HResult").GetString() != "0x0") throw new ArchiveReviewException("Localização do PST incompleta.");
            string? id = result.GetProperty("StoreId").GetString();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }
        finally { if (File.Exists(request)) File.Delete(request); }
    }
}
