using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;

internal static class PstReopenCheck
{
    internal static int Execute(string runDirectory)
    {
        object? app = null, session = null, store = null, root = null;
        DurableJournal? journal = null;
        string? pstPath = null, evidence = null;
        bool detached = false, reattached = false;
        var result = new ReopenResult();
        try
        {
            string run = Path.GetFullPath(runDirectory);
            var report = JsonSerializer.Deserialize<PilotReport>(File.ReadAllText(Path.Combine(run, "report.json")))
                ?? throw new InvalidDataException("Relatório inválido.");
            pstPath = ValidateReport(report);
            var audit = PilotAudit.Read(run);
            if (audit.Errors.Count != 0 || audit.ImportAmbiguous.Length != 0 || audit.CopyCreationAmbiguous.Length != 0 ||
                audit.TemporaryCopiesPossiblyRemaining.Length != 0 || audit.MetadataRestoreAmbiguous.Length != 0 || audit.DestinationRelocations.Count != 0)
                throw new InvalidOperationException("A execução exige reconciliação prévia de identidades/operações.");
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            store = FindStore(session, pstPath) ?? throw new InvalidOperationException("PST do piloto não está anexado.");
            string initialStoreId = (string)((dynamic)store).StoreID;
            if (initialStoreId != report.TargetStoreId) throw new InvalidOperationException("Identidade do PST mudou; reconcilie antes de continuar.");
            object? defaultStore = null;
            try
            {
                defaultStore = ((dynamic)session).DefaultStore;
                if ((string)((dynamic)defaultStore).StoreID == initialStoreId)
                    throw new InvalidOperationException("O armazenamento padrão não pode ser desanexado por este teste.");
            }
            finally { Release(defaultStore); }
            CheckOpenWindows(app!, initialStoreId);
            evidence = Path.Combine(run, "reopen-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(evidence);
            result.PstPath = pstPath;
            journal = new DurableJournal(Path.Combine(evidence, "journal.jsonl"));
            journal.Write(new { State = "detach-started", PstPath = pstPath, PreviousStoreId = initialStoreId });
            root = ((dynamic)store).GetRootFolder();
            detached = true;
            ((dynamic)session).RemoveStore(root); // Detach only; never deletes a file or item.
            Release(root); root = null;
            Release(store); store = null;
            journal.Write(new { State = "detached" });
            // An exclusive disk read establishes that this is more than a UI-only
            // reattachment. If Outlook keeps a handle, record the limitation.
            for (int attempt = 0; attempt < 10; attempt++)
            {
                try
                {
                    using var input = new FileStream(pstPath, FileMode.Open, FileAccess.Read, FileShare.None);
                    result.PstBytesOnDisk = input.Length;
                    result.PstSha256WhileDetached = Convert.ToHexString(SHA256.HashData(input));
                    result.ExclusiveDiskReadConfirmed = true;
                    break;
                }
                catch (IOException) { if (attempt < 9) Thread.Sleep(200); }
            }
            journal.Write(new { State = "disk-probe", result.ExclusiveDiskReadConfirmed, result.PstBytesOnDisk, result.PstSha256WhileDetached });
            using (var guard = new FileStream(pstPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                ((dynamic)session).AddStoreEx(pstPath, 2); // Guard forbids path deletion/replacement.
            reattached = true;
            store = FindStore(session, pstPath) ?? throw new InvalidOperationException("PST não localizado após reabertura.");
            string currentStoreId = (string)((dynamic)store).StoreID;
            result.ReopenedStoreId = currentStoreId;
            journal.Write(new { State = "reattached", StoreId = currentStoreId });
            int index = 0;
            foreach (var item in report.Items)
            {
                index++;
                object? copy = null, parent = null;
                var row = new ReopenedCopy { Index = index, DestinationId = item.DestinationId! };
                result.Items.Add(row);
                try
                {
                    copy = ((dynamic)session).GetItemFromID(item.DestinationId, currentStoreId);
                    parent = ((dynamic)copy).Parent;
                    if ((string)((dynamic)parent).StoreID != currentStoreId || (string)((dynamic)parent).EntryID != report.TargetFolderId)
                        throw new InvalidOperationException("Cópia não pertence à pasta esperada.");
                    string directory = Path.Combine(evidence, $"item-{index:D3}");
                    Directory.CreateDirectory(directory);
                    string msg = Path.Combine(directory, "destination-reopened.msg");
                    ((dynamic)copy).SaveAs(msg, 9);
                    PilotCommand.Snapshot(copy, directory, msg);
                    string beforePath = Path.Combine(run, $"item-{index:D3}", "destination", "field-hashes.json");
                    var before = ReadHashes(beforePath);
                    var after = ReadHashes(Path.Combine(directory, "field-hashes.json"));
                    row.DifferentFields = before.Keys.Union(after.Keys).Where(key => before.GetValueOrDefault(key) != after.GetValueOrDefault(key)).Order(StringComparer.Ordinal).ToArray();
                    row.Status = row.DifferentFields.Length == 0 ? "reopened-snapshot-equal" : "review-required";
                    journal.Write(row);
                }
                catch (Exception ex) { row.Error = ex.GetType().Name; row.Status = "review-required"; journal.Write(row); }
                finally { Release(parent); Release(copy); }
            }
            result.AllReopenedCopiesMatchSnapshots = result.Items.Count == report.Items.Count &&
                result.Items.All(x => x.Status == "reopened-snapshot-equal");
            result.Status = result.AllReopenedCopiesMatchSnapshots && result.ExclusiveDiskReadConfirmed
                ? "reopen-persistence-check-passed" : "review-required";
            journal.Write(new { State = result.Status, result.AllReopenedCopiesMatchSnapshots, result.ExclusiveDiskReadConfirmed });
            SaveResult(evidence, result);
            Console.WriteLine($"{result.Status}: {result.Items.Count(x => x.Status == "reopened-snapshot-equal")}/{report.Items.Count} cópias iguais após reabertura. Leitura exclusiva do disco: {result.ExclusiveDiskReadConfirmed}.\nEvidência: {evidence}");
            return result.Status == "reopen-persistence-check-passed" ? 0 : 3;
        }
        catch (Exception ex)
        {
            result.Status = "review-required";
            result.Error = ex.GetType().Name + ": " + ex.Message;
            Console.Error.WriteLine("Verificação de reabertura interrompida: " + ex.Message);
            if (evidence is not null)
            {
                try { SaveResult(evidence, result); } catch (IOException) { }
            }
            return 1;
        }
        finally
        {
            Release(root); Release(store);
            if (detached && !reattached && session is not null && pstPath is not null && File.Exists(pstPath))
            {
                try
                {
                    object? existing = null;
                    try
                    {
                        existing = FindStore(session, pstPath);
                        if (existing is null)
                        {
                            using var guard = new FileStream(pstPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            ((dynamic)session).AddStoreEx(pstPath, 2);
                            Console.WriteLine("PST original reanexado após a interrupção.");
                        }
                    }
                    finally { Release(existing); }
                }
                catch (Exception ex) { Console.Error.WriteLine("Reanexe manualmente o PST original: " + pstPath + ". " + ex.GetType().Name); }
            }
            journal?.Dispose(); Release(session); Release(app);
        }
    }

    internal static string ValidateReport(PilotReport report)
    {
        if (string.IsNullOrEmpty(report.PstPath) || !Path.IsPathFullyQualified(report.PstPath) ||
            !string.Equals(Path.GetExtension(report.PstPath), ".pst", StringComparison.OrdinalIgnoreCase) ||
            !File.Exists(report.PstPath)) throw new InvalidDataException("PST existente com caminho absoluto é obrigatório.");
        if (string.IsNullOrEmpty(report.TargetStoreId) || report.TargetStoreId == report.SourceStoreId ||
            string.IsNullOrEmpty(report.TargetFolderId)) throw new InvalidDataException("Destino não pode ser a origem e deve possuir identidades registradas.");
        if (report.Items.Count is < 1 or > 25 || report.Items.Any(x => string.IsNullOrEmpty(x.DestinationId) || x.DestinationId == x.SourceId) ||
            report.Items.Select(x => x.DestinationId).Distinct(StringComparer.Ordinal).Count() != report.Items.Count)
            throw new InvalidDataException("O teste exige até 25 cópias registradas com identidades distintas.");
        return Path.GetFullPath(report.PstPath);
    }
    private static void CheckOpenWindows(object app, string storeId)
    {
        object? inspectors = null, explorers = null;
        try
        {
            inspectors = ((dynamic)app).Inspectors;
            for (int i = 1; i <= (int)((dynamic)inspectors).Count; i++)
            {
                object? inspector = null, item = null, parent = null;
                try
                {
                    inspector = ((dynamic)inspectors)[i]; item = ((dynamic)inspector).CurrentItem;
                    parent = ((dynamic)item).Parent;
                    if ((string)((dynamic)parent).StoreID == storeId)
                        throw new InvalidOperationException("Há mensagem do PST do piloto aberta em uma janela; o teste não a fechará.");
                }
                finally { Release(parent); Release(item); Release(inspector); }
            }
            explorers = ((dynamic)app).Explorers;
            for (int i = 1; i <= (int)((dynamic)explorers).Count; i++)
            {
                object? explorer = null, folder = null;
                try
                {
                    explorer = ((dynamic)explorers)[i]; folder = ((dynamic)explorer).CurrentFolder;
                    if ((string)((dynamic)folder).StoreID == storeId)
                        throw new InvalidOperationException("Uma janela está exibindo o PST do piloto; selecione outra caixa antes do teste.");
                }
                finally { Release(folder); Release(explorer); }
            }
        }
        finally { Release(explorers); Release(inspectors); }
    }
    private static object? FindStore(object session, string path)
    {
        object? stores = null;
        try
        {
            stores = ((dynamic)session).Stores;
            for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
            {
                object? candidate = null;
                try
                {
                    candidate = ((dynamic)stores)[i];
                    string candidatePath = (string)((dynamic)candidate).FilePath;
                    if (!string.IsNullOrWhiteSpace(candidatePath) && string.Equals(Path.GetFullPath(candidatePath), path, StringComparison.OrdinalIgnoreCase))
                    { object found = candidate; candidate = null; return found; }
                }
                catch (Exception ex) when (ex is COMException or ArgumentException) { }
                finally { Release(candidate); }
            }
            return null;
        }
        finally { Release(stores); }
    }
    internal static void SelfTest()
    {
        string file = Path.Combine(Path.GetTempPath(), "outlook-reopen-guard-" + Guid.NewGuid().ToString("N") + ".pst");
        File.WriteAllText(file, "fixture: report validation only");
        try
        {
            var report = new PilotReport { PstPath = file, SourceStoreId = "source", TargetStoreId = "target", TargetFolderId = "folder",
                Items = [new PilotItemResult { SourceId = "original", DestinationId = "copy" }] };
            ValidateReport(report);
            report.TargetStoreId = "source";
            MustReject(report);
            report.TargetStoreId = "target";
            report.Items[0].DestinationId = "original";
            MustReject(report);
            report.Items[0].DestinationId = "copy";
            report.Items.Add(new PilotItemResult { SourceId = "other", DestinationId = "copy" });
            MustReject(report);
            report.Items.RemoveAt(1);
            File.Delete(file);
            MustReject(report);
            Console.WriteLine("Reabertura: destino igual à origem, IDs repetidos/originais e arquivo ausente bloqueados antes do COM.");
        }
        finally { File.Delete(file); }
    }
    private static void MustReject(PilotReport report)
    {
        try { ValidateReport(report); }
        catch (InvalidDataException) { return; }
        throw new Exception("Relatório inválido aceito para reabertura.");
    }
    private static Dictionary<string, string> ReadHashes(string path)
    {
        var hashes = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path));
        return hashes is { Count: > 0 } ? hashes : throw new InvalidDataException("Evidência de comparação vazia.");
    }
    private static void SaveResult(string directory, ReopenResult result)
    {
        string temp = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".partial.json");
        using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        { JsonSerializer.Serialize(stream, result, new JsonSerializerOptions { WriteIndented = true }); stream.Flush(true); }
        File.Move(temp, Path.Combine(directory, "report.json"), false);
    }
    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? instance);
}
internal sealed class ReopenResult
{
    public string Status { get; set; } = "review-required";
    public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
    public string? Error { get; set; }
    public string? PstPath { get; set; }
    public string? ReopenedStoreId { get; set; }
    public bool ExclusiveDiskReadConfirmed { get; set; }
    public long? PstBytesOnDisk { get; set; }
    public string? PstSha256WhileDetached { get; set; }
    public bool AllReopenedCopiesMatchSnapshots { get; set; }
    public bool RemovalAllowed { get; } = false;
    public string Scope { get; } = "Desanexar/reabrir PST do piloto na mesma sessão; não reinicia Outlook/Windows nem certifica completude da origem.";
    public List<ReopenedCopy> Items { get; set; } = new();
}
internal sealed class ReopenedCopy
{
    public int Index { get; set; }
    public string DestinationId { get; set; } = "";
    public string Status { get; set; } = "review-required";
    public string? Error { get; set; }
    public string[] DifferentFields { get; set; } = [];
}
