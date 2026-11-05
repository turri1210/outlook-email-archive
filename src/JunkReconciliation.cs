using System.Runtime.InteropServices;
using System.Text.Json;

// Narrow recovery experiment: relocate only an already recorded destination copy
// inside its PST to the equivalent default Junk folder, then compare again.
internal static class JunkReconciliation
{
    internal static int Execute(string runDirectory)
    {
        object? app = null, session = null, sourceFolder = null, sourceStore = null, sourceJunk = null;
        object? destinationFolder = null, destinationStore = null, destinationJunk = null;
        var results = new List<object>();
        try
        {
            string run = Path.GetFullPath(runDirectory);
            var report = JsonSerializer.Deserialize<PilotReport>(File.ReadAllText(Path.Combine(run, "report.json")))
                ?? throw new InvalidDataException("Relatório inválido.");
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            sourceFolder = ((dynamic)session).GetFolderFromID(report.SourceFolderId, report.SourceStoreId);
            sourceStore = ((dynamic)sourceFolder).Store;
            sourceJunk = ((dynamic)sourceStore).GetDefaultFolder(23);
            if ((string)((dynamic)sourceJunk).EntryID != report.SourceFolderId)
                throw new InvalidOperationException("Origem não é a pasta padrão Junk Email.");
            destinationFolder = ((dynamic)session).GetFolderFromID(report.TargetFolderId, report.TargetStoreId);
            destinationStore = ((dynamic)destinationFolder).Store;
            if (string.Equals(report.SourceStoreId, report.TargetStoreId, StringComparison.Ordinal) ||
                !string.Equals(Path.GetFullPath((string)((dynamic)destinationStore).FilePath), Path.GetFullPath(report.PstPath!), StringComparison.OrdinalIgnoreCase) ||
                Path.GetExtension(report.PstPath) != ".pst") throw new InvalidOperationException("Destino não é o PST do relatório.");
            destinationJunk = ((dynamic)destinationStore).GetDefaultFolder(23);
            if (destinationJunk is null) throw new InvalidOperationException("PST não possui pasta Junk padrão; nenhuma cópia será movida.");
            string evidence = Path.Combine(run, "reconciliation-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(evidence);
            using var journal = new DurableJournal(Path.Combine(evidence, "journal.jsonl"));
            journal.Write(new { State = "reconciliation-started", TargetFolderId = (string)((dynamic)destinationJunk).EntryID });
            int index = 0;
            bool allVerified = true;
            foreach (var item in report.Items.Where(x => !string.IsNullOrEmpty(x.DestinationId)))
            {
                index++;
                object? original = null, copy = null, moved = null, parent = null, reopened = null;
                try
                {
                    original = ((dynamic)session).GetItemFromID(item.SourceId, report.SourceStoreId);
                    copy = ((dynamic)session).GetItemFromID(item.DestinationId, report.TargetStoreId);
                    parent = ((dynamic)copy).Parent;
                    if ((string)((dynamic)parent).StoreID != report.TargetStoreId || item.SourceId == item.DestinationId)
                        throw new InvalidOperationException("Cópia não está no armazenamento esperado.");
                    journal.Write(new { State = "destination-relocation-started", item.SourceId, item.DestinationId });
                    moved = ((dynamic)copy).Move(destinationJunk);
                    string newId = (string)((dynamic)moved).EntryID;
                    journal.Write(new { State = "destination-relocated", PreviousId = item.DestinationId, DestinationId = newId });
                    reopened = ((dynamic)session).GetItemFromID(newId, report.TargetStoreId);
                    string sourceDir = Path.Combine(evidence, $"source-{index:D3}");
                    string targetDir = Path.Combine(evidence, $"destination-{index:D3}");
                    Directory.CreateDirectory(sourceDir); Directory.CreateDirectory(targetDir);
                    string originalMsg = Path.Combine(sourceDir, "message.msg"), copyMsg = Path.Combine(targetDir, "message.msg");
                    ((dynamic)original).SaveAs(originalMsg, 9); ((dynamic)reopened).SaveAs(copyMsg, 9);
                    string left = PilotCommand.Snapshot(original, sourceDir, originalMsg);
                    string right = PilotCommand.Snapshot(reopened, targetDir, copyMsg);
                    var leftFields = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(sourceDir, "field-hashes.json")))!;
                    var rightFields = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(targetDir, "field-hashes.json")))!;
                    var differences = leftFields.Keys.Union(rightFields.Keys).Where(key => leftFields.GetValueOrDefault(key) != rightFields.GetValueOrDefault(key)).ToArray();
                    allVerified &= left == right;
                    var result = new { item.SourceId, PreviousDestinationId = item.DestinationId, DestinationId = newId,
                        Status = left == right ? "verified-selected-fields" : "review-required", Differences = differences };
                    results.Add(result); journal.Write(result);
                    Console.WriteLine($"Cópia {index}: {result.Status}; divergências: {string.Join(", ", differences)}");
                }
                finally { Release(reopened); Release(moved); Release(parent); Release(copy); Release(original); }
            }
            if (index == 0) throw new InvalidOperationException("Nenhuma cópia registrada para reconciliar.");
            File.WriteAllText(Path.Combine(evidence, "report.json"), JsonSerializer.Serialize(new
            { Status = allVerified ? "reconciled-selected-fields" : "review-required", RemovalEnabled = false, Results = results }, new JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine("Evidência: " + evidence);
            return allVerified ? 0 : 3;
        }
        catch (Exception ex) { Console.Error.WriteLine("Reconciliação interrompida: " + ex.Message); return 1; }
        finally
        {
            Release(destinationJunk); Release(destinationStore); Release(destinationFolder);
            Release(sourceJunk); Release(sourceStore); Release(sourceFolder); Release(session); Release(app);
        }
    }
    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);
}
