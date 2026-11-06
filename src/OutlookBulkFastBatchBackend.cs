using System.Diagnostics;
using System.Text;
using System.Text.Json;

internal sealed partial class OutlookArchiveBackend : IArchiveBulkFastBatchBackend
{
    private sealed record BulkIdentity(FastMessageIdentity? Identity, bool CanBatch, string? Error);
    private readonly Dictionary<string, BulkIdentity> bulkSourceIdentities = new(StringComparer.Ordinal);
    private readonly Dictionary<string, BulkIdentity> bulkTargetIdentities = new(StringComparer.Ordinal);
    private readonly HashSet<string> nativeUnavailableSourceStores=new(StringComparer.OrdinalIgnoreCase);
    private static string SourceBulkKey(ArchiveItem row) => row.SourceStoreId + "|" + row.SourceEntryId + "|" + row.SourceFolderId + "|" + row.SourceRecordKey;
    private static string TargetBulkKey(DestinationIdentity identity) => identity.StoreId + "|" + identity.FolderId + "|" + identity.EntryId;
    private static FastMessageIdentity RequireBulkIdentity(BulkIdentity result)
    {
        if (result.Error is not null || result.Identity is null)
            throw new ArchiveReviewException("Identificação nativa da mensagem pendente: " + (result.Error ?? "resposta ausente"));
        return result.Identity;
    }
    public void ClearFastIdentityCache() { bulkSourceIdentities.Clear(); bulkTargetIdentities.Clear(); }
    public void RefreshSourceIdentities(IReadOnlyList<ArchiveItem> rows)
    {
        foreach (var group in rows.GroupBy(x => x.SourceStoreId))
        foreach (var chunk in group.Chunk(100))
        {
            if(nativeUnavailableSourceStores.Contains(group.Key))continue;
            Dictionary<string,BulkIdentity> result;
            try{result=ReadBulkIdentities(group.Key, chunk.Select(x => (x.SourceEntryId, x.SourceFolderId, x.SourceRecordKey)).ToArray());}
            catch(NativeStoreAccessException)
            {
                // The helper is read-only and failed before opening a message.
                // Prove that Outlook itself can open this exact source first.
                object? store=null,folder=null;
                try
                {
                    store=((dynamic)session!).GetStoreFromID(group.Key);folder=((dynamic)store).GetRootFolder();
                    if(!SameIdentity((string)((dynamic)folder).StoreID,group.Key))throw new ArchiveReviewException("O Outlook retornou outro armazenamento de origem.");
                    _=(string)((dynamic)folder).EntryID;
                }
                finally{Release(folder);Release(store);}
                nativeUnavailableSourceStores.Add(group.Key);
                Console.WriteLine("Cópia: esta origem está acessível pelo Outlook, mas não pelo componente nativo (0x8004011D). As mensagens sem lote iniciado usarão o transporte individual por MSG; validação e proteção dos originais permanecem obrigatórias.");
                continue;
            }
            foreach (var row in chunk) bulkSourceIdentities[SourceBulkKey(row)] = result.GetValueOrDefault(row.SourceEntryId)
                ?? new(null, false, "Mensagem ausente da resposta nativa.");
        }
    }
    public void RefreshTargetIdentities(ArchiveVolume volume, IReadOnlyList<DestinationIdentity> candidates)
    {
        if (candidates.Any(x => x.StoreId != volume.StoreId)) throw new ArchiveReviewException("Candidato nativo fora do PST registrado.");
        if (candidates.All(x => bulkTargetIdentities.ContainsKey(TargetBulkKey(x)))) return;
        foreach (var chunk in candidates.Chunk(100))
        {
            var result = ReadBulkIdentities(volume.StoreId!, chunk.Select(x => (x.EntryId, x.FolderId, (string?)null)).ToArray());
            foreach (var identity in chunk) bulkTargetIdentities[TargetBulkKey(identity)] = result.GetValueOrDefault(identity.EntryId)
                ?? new(null, false, "Cópia ausente da resposta nativa.");
        }
    }
    private List<DestinationIdentity> ReadBulkBatchCandidates(ArchiveVolume volume, ArchiveBatch batch)
    {
        object? folder = null;
        try { folder = OpenBatchFolder(volume, batch); }
        finally { Release(folder); }
        var captured = ReadBulkIdentities(volume.StoreId!, [], batch.FolderId!, batch.Items.Count);
        var identities = captured.Keys.Select(id => new DestinationIdentity(volume.StoreId!, batch.FolderId!, id)).ToList();
        foreach (var identity in identities) bulkTargetIdentities[TargetBulkKey(identity)] = captured[identity.EntryId];
        return identities;
    }
    private bool HasFreshBulkScope(IReadOnlyList<ArchiveItem> rows) => rows.All(row =>
        bulkSourceIdentities.TryGetValue(SourceBulkKey(row), out var result) && result.Error is null && result.CanBatch &&
        result.Identity is not null && result.Identity.Hash == row.FastIdentityHash);
    private static Dictionary<string, BulkIdentity> ReadBulkIdentities(string storeId, IReadOnlyList<(string EntryId, string FolderId, string? RecordKey)> rows, string? folderId = null, int maximum = 100)
    {
        var result = new Dictionary<string, BulkIdentity>(StringComparer.Ordinal);
        if (rows.Count == 0 && folderId is null) return result;
        string executable = Path.Combine(AppContext.BaseDirectory, "native-identity.exe");
        if (!File.Exists(executable)) throw new ArchiveReviewException("Componente nativo de identificação ausente.");
        string handoff = Path.Combine(Path.GetTempPath(), "outlook-identity-" + Guid.NewGuid().ToString("N") + ".txt");
        try
        {
            using (var stream = new FileStream(handoff, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.WriteLine(folderId is null ? "readonly-identity-v1" : "readonly-folder-v1"); writer.WriteLine(storeId);
                if (folderId is not null) { writer.WriteLine(folderId); writer.WriteLine(maximum); }
                foreach (var row in rows) writer.WriteLine(row.EntryId + "|" + row.FolderId + "|" + row.RecordKey);
            }
            using var process = new Process { StartInfo = new ProcessStartInfo(executable)
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8 } };
            process.StartInfo.ArgumentList.Add(handoff);
            if (!process.Start()) throw new IOException("Não foi possível iniciar a identificação nativa.");
            var stdout = process.StandardOutput.ReadToEndAsync(); var stderr = process.StandardError.ReadToEndAsync();
            process.WaitForExit(); string json = stdout.GetAwaiter().GetResult(), errors = stderr.GetAwaiter().GetResult();
            if (process.ExitCode != 0) throw NativeStoreAccess.Failure(json,errors);
            using var document = JsonDocument.Parse(json); var root = document.RootElement;
            if (root.GetProperty("Stage").GetString() != "complete" || root.GetProperty("HResult").GetString() != "0x0")
                throw new ArchiveReviewException("Identificação nativa incompleta.");
            var expected = rows.Select(x => x.EntryId).ToHashSet(StringComparer.Ordinal);
            foreach (var row in root.GetProperty("Rows").EnumerateArray())
            {
                string entryId = row.GetProperty("EntryId").GetString() ?? throw new ArchiveReviewException("Identidade nativa ausente.");
                if ((folderId is null && !expected.Contains(entryId)) || result.ContainsKey(entryId)) throw new ArchiveReviewException("Resposta nativa contém identidade extra ou repetida.");
                if (row.TryGetProperty("Error", out var error)) result.Add(entryId, new(null, false, error.GetString()));
                else
                {
                    var fields = row.GetProperty("Fields").Deserialize<Dictionary<string, string>>() ?? throw new ArchiveReviewException("Metadados nativos ausentes.");
                    result.Add(entryId, new(new(fields, row.GetProperty("ChangeKey").GetString()), row.GetProperty("CanBatch").GetBoolean(), null));
                }
            }
            if (folderId is null ? result.Count != rows.Count : result.Count > maximum) throw new ArchiveReviewException("Resposta nativa não cobre o escopo solicitado.");
            return result;
        }
        finally { if (File.Exists(handoff)) File.Delete(handoff); }
    }
}
