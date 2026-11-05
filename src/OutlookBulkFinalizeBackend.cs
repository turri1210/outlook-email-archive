using System.Diagnostics;
using System.Text;
using System.Text.Json;

internal interface IArchiveBulkFinalizeBackend
{
    Dictionary<string, DestinationIdentity> FinalizeBatchCopies(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> rows);
}
internal sealed partial class OutlookArchiveBackend : IArchiveBulkFinalizeBackend
{
    public Dictionary<string, DestinationIdentity> FinalizeBatchCopies(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> rows)
    {
        ValidateBatch(volume, batch); AttachVolume(volume); AssertBatchIdle(batch);
        if (!batch.FastCopy || rows.Count is < 1 or > 100 || rows.Any(x => !batch.Items.Contains(x.Key) || !batch.Matches.ContainsKey(x.Key) || x.JobId != batch.JobId ||
            !x.RelativePath.SequenceEqual(batch.RelativePath)) || rows.Select(x => x.Key).Distinct().Count() != rows.Count)
            throw new ArchiveReviewException("Escopo inválido para finalização em lote.");
        string executable = Path.Combine(AppContext.BaseDirectory, "native-finalize.exe");
        if (!File.Exists(executable)) throw new ArchiveReviewException("Componente nativo de finalização ausente.");
        object? staging = null, final = null;
        string finalId;
        try
        {
            staging = OpenBatchFolder(volume, batch);
            final = TargetFolder(volume, batch.RelativePath, true) ?? throw new ArchiveReviewException("Pasta final indisponível.");
            if (!BatchSame((string)((dynamic)final).StoreID, volume.StoreId!)) throw new ArchiveReviewException("Pasta final fora do PST registrado.");
            finalId = (string)((dynamic)final).EntryID;
        }
        finally { Release(final); Release(staging); }
        using var launch = new FileStream(Path.Combine(batch.DirectoryPath, "launch.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        string handoff = Path.Combine(batch.DirectoryPath, "finalize-" + Guid.NewGuid().ToString("N") + ".txt");
        using (var stream = new FileStream(handoff, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            writer.WriteLine("pst-finalize-only-v1"); writer.WriteLine(Path.Combine(batch.DirectoryPath, "worker.lock"));
            writer.WriteLine(volume.StoreId); writer.WriteLine(batch.FolderId); writer.WriteLine(finalId);
            foreach (var row in rows) writer.WriteLine(row.Key + "|" + batch.Matches[row.Key].EntryId + "|" + row.Token);
            writer.Flush(); stream.Flush(true);
        }
        using var process = new Process { StartInfo = new ProcessStartInfo(executable)
        { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true, StandardOutputEncoding = Encoding.UTF8 } };
        process.StartInfo.ArgumentList.Add(handoff);
        if (!process.Start()) throw new IOException("Não foi possível iniciar a finalização nativa.");
        var output = process.StandardOutput.ReadToEndAsync(); var errors = process.StandardError.ReadToEndAsync(); process.WaitForExit();
        string json = output.GetAwaiter().GetResult(), stderr = errors.GetAwaiter().GetResult();
        File.WriteAllText(Path.Combine(batch.DirectoryPath, "finalize-output.json"), json);
        File.WriteAllText(Path.Combine(batch.DirectoryPath, "finalize-errors.txt"), stderr);
        tokenIndexes.Remove(volume.PstPath); batchRecoveryIndexes.Remove(batch.Id);
        using var document = JsonDocument.Parse(json); var result = document.RootElement;
        if (process.ExitCode != 0 || result.GetProperty("Stage").GetString() != "complete" || result.GetProperty("HResult").GetString() != "0x0")
            throw new ArchiveReviewException("Finalização interna do PST pendente: " + result.GetProperty("Error").GetString() + ". Retome a mesma execução; a cópia original não será repetida.");
        var matches = result.GetProperty("Matches").Deserialize<Dictionary<string, string>>() ?? throw new ArchiveReviewException("Mapa final ausente.");
        if (matches.Count != rows.Count || rows.Any(x => !matches.ContainsKey(x.Key)) || matches.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != rows.Count ||
            matches.Values.Any(x => x.Length == 0 || x.Length % 2 != 0 || x.Any(c => !Uri.IsHexDigit(c))))
            throw new ArchiveReviewException("Mapa da finalização não corresponde ao escopo registrado.");
        return matches.ToDictionary(x => x.Key, x => new DestinationIdentity(volume.StoreId!, finalId, x.Value), StringComparer.Ordinal);
    }
}
