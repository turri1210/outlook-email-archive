using System.Runtime.InteropServices;
using System.Text.Json;

// Metadata-only scan of the selected folder's direct Items collection.
// Never enumerates other folders, reads bodies/subjects/attachment bytes,
// launches Outlook, or changes source items.
internal static class FolderCoverage
{
    private const int MaximumItems = 10000;

    internal static int Execute(string[] args)
    {
        if (args.Length != 3 || args.Any(string.IsNullOrWhiteSpace))
        {
            Console.Error.WriteLine("Uso: folder-coverage <store-id> <folder-entry-id> <arquivo.json>");
            return 2;
        }
        object? app = null, session = null, folder = null, items = null;
        var report = new CoverageReport { SourceStoreId = args[0], SourceFolderId = args[1] };
        string? output = null;
        try
        {
            output = Path.GetFullPath(args[2]);
            if (File.Exists(output) || Directory.Exists(output))
                throw new IOException("O destino já existe; sobrescrita proibida.");
            if (!Directory.Exists(Path.GetDirectoryName(output)))
                throw new DirectoryNotFoundException("A pasta do relatório deve existir.");

            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            folder = ((dynamic)session).GetFolderFromID(args[1], args[0]);
            if (!string.Equals((string)((dynamic)folder).StoreID, args[0], StringComparison.Ordinal) ||
                !string.Equals((string)((dynamic)folder).EntryID, args[1], StringComparison.Ordinal))
                throw new InvalidOperationException("A pasta retornada não corresponde à origem selecionada.");

            object? accessor = null;
            try
            {
                accessor = ((dynamic)folder).PropertyAccessor;
                if ((int)((dynamic)accessor).GetProperty("http://schemas.microsoft.com/mapi/proptag/0x36010003") == 2)
                    throw new InvalidOperationException("Pasta de pesquisa não é aceita como origem direta.");
            }
            finally { Release(accessor); }

            items = ((dynamic)folder).Items;
            report.CountBefore = (int)((dynamic)items).Count;
            report.LimitReached = report.CountBefore > MaximumItems;
            int plannedCount = Math.Min(report.CountBefore.Value, MaximumItems);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            for (int index = 1; index <= plannedCount; index++)
            {
                object? item = null;
                var row = new CoverageItem { Index = index };
                report.Items.Add(row);
                try
                {
                    item = ((dynamic)items)[index];
                    dynamic current = item;
                    ReadValue(report, index, "entry-id", () => { row.EntryId = (string)current.EntryID; });
                    if (row.EntryId is not null && string.IsNullOrWhiteSpace(row.EntryId))
                    {
                        row.EntryId = null;
                        report.Failures.Add(new(index, "entry-id", "empty-entry-id", null));
                    }
                    if (row.EntryId is not null)
                    {
                        if (!ids.Add(row.EntryId))
                            report.Failures.Add(new(index, "entry-id", "duplicate-observed-id", null));
                    }
                    ReadValue(report, index, "class", () => { row.Class = (int)current.Class; });
                    ReadValue(report, index, "message-class", () => { row.MessageClass = (string)current.MessageClass; });
                    ReadValue(report, index, "size", () => { row.SizeBytes = (int)current.Size; });
                    if (row.SizeBytes < 0)
                    {
                        row.SizeBytes = null;
                        report.Failures.Add(new(index, "size", "negative-item-size", null));
                    }
                    // BodyFormat belongs to MailItem. Its integer value is metadata;
                    // do not request Body, HTMLBody, RTFBody or PropertyAccessor bodies.
                    if (row.Class == 43)
                    {
                        ReadValue(report, index, "body-format", () => { row.BodyFormat = (int)current.BodyFormat; });
                        row.BodyFormatStatus = row.BodyFormat.HasValue ? "observed" : "unavailable";
                    }
                    else row.BodyFormatStatus = row.Class.HasValue ? "not-applicable-to-item-class" : "unknown-item-class";

                    object? attachments = null;
                    try
                    {
                        attachments = current.Attachments;
                        row.AttachmentCount = (int)((dynamic)attachments).Count;
                        if (row.AttachmentCount < 0)
                        {
                            row.AttachmentCount = null;
                            throw new InvalidDataException("Contagem de anexos inválida.");
                        }
                        for (int attachmentIndex = 1; attachmentIndex <= row.AttachmentCount.Value; attachmentIndex++)
                        {
                            object? attachment = null;
                            try
                            {
                                attachment = ((dynamic)attachments)[attachmentIndex];
                                row.AttachmentTypes.Add((int)((dynamic)attachment).Type);
                            }
                            catch (Exception ex)
                            { report.Failures.Add(Failure(index, $"attachment-type:{attachmentIndex}", ex)); }
                            finally { Release(attachment); }
                        }
                    }
                    catch (Exception ex) { report.Failures.Add(Failure(index, "attachments-metadata", ex)); }
                    finally { Release(attachments); }
                }
                catch (Exception ex) { report.Failures.Add(Failure(index, "item", ex)); }
                finally { Release(item); }
            }
            try { report.CountAfter = (int)((dynamic)items).Count; }
            catch (Exception ex) { report.Failures.Add(Failure(null, "count-after", ex)); }
            report.UniqueEntryIds = ids.Count;
            report.ObservedSizeBytesSum = report.Items.Where(x => x.SizeBytes.HasValue).Sum(x => (long)x.SizeBytes!.Value);
            report.ItemsWithKnownSize = report.Items.Count(x => x.SizeBytes.HasValue);
            report.ItemsWithAttachments = report.Items.Count(x => x.AttachmentCount > 0);
            report.ItemsWithUnknownAttachmentCount = report.Items.Count(x => !x.AttachmentCount.HasValue);
            report.AttachmentCountObserved = report.Items.Where(x => x.AttachmentCount.HasValue).Sum(x => (long)x.AttachmentCount!.Value);
            report.AttachmentTypeCounts = report.Items.SelectMany(x => x.AttachmentTypes)
                .GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
            report.MessageClassCounts = report.Items.Where(x => x.MessageClass is not null)
                .GroupBy(x => x.MessageClass!, StringComparer.Ordinal).ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);
            report.BodyFormatCounts = report.Items.Where(x => x.BodyFormat.HasValue)
                .GroupBy(x => x.BodyFormat!.Value).ToDictionary(g => g.Key, g => g.Count());
            report.Status = report.Failures.Count == 0 && !report.LimitReached &&
                report.CountBefore == report.CountAfter && report.UniqueEntryIds == plannedCount
                ? "metadata-observed" : "review-required";
            report.FinishedUtc = DateTimeOffset.UtcNow;
            WriteAtomically(output, report);
            Console.WriteLine($"Cobertura: {report.Items.Count} itens diretos; {report.Failures.Count} falhas; " +
                $"{report.ItemsWithAttachments} itens com anexos. Estado: {report.Status}.");
            return report.Status == "metadata-observed" ? 0 : 3;
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
        {
            Console.Error.WriteLine("Outlook clássico não acessível por COM. Abra-o, aguarde o perfil carregar e alterne para outra janela.");
            return 1;
        }
        catch (Exception ex)
        {
            // Avoid exception text that could include data supplied by a mail item.
            Console.Error.WriteLine($"Cobertura interrompida: {ex.GetType().Name} (0x{ex.HResult:X8}).");
            return 1;
        }
        finally { Release(items); Release(folder); Release(session); Release(app); }
    }

    private static void ReadValue(CoverageReport report, int index, string stage, Action read)
    {
        try { read(); }
        catch (Exception ex) { report.Failures.Add(Failure(index, stage, ex)); }
    }
    private static CoverageFailure Failure(int? index, string stage, Exception exception) =>
        new(index, stage, exception.GetType().Name, $"0x{exception.HResult:X8}");

    private static void WriteAtomically(string output, CoverageReport report)
    {
        string temporary = Path.Combine(Path.GetDirectoryName(output)!, $".{Path.GetFileName(output)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                JsonSerializer.Serialize(stream, report, new JsonSerializerOptions { WriteIndented = true });
                stream.Flush(true);
            }
            File.Move(temporary, output, overwrite: false);
        }
        finally
        {
            try { File.Delete(temporary); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);

    private sealed record CoverageFailure(int? Index, string Stage, string ErrorTypeOrCode, string? HResult);
    private sealed class CoverageItem
    {
        public int Index { get; set; }
        public string? EntryId { get; set; }
        public int? Class { get; set; }
        public string? MessageClass { get; set; }
        public int? SizeBytes { get; set; }
        public int? BodyFormat { get; set; }
        public string BodyFormatStatus { get; set; } = "unavailable";
        public int? AttachmentCount { get; set; }
        public List<int> AttachmentTypes { get; set; } = new();
        public bool AttachmentTypesComplete => AttachmentCount.HasValue && AttachmentTypes.Count == AttachmentCount.Value;
    }
    private sealed class CoverageReport
    {
        public int SchemaVersion { get; set; } = 1;
        public string Status { get; set; } = "started";
        public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
        public DateTimeOffset? FinishedUtc { get; set; }
        public string SourceStoreId { get; set; } = "";
        public string SourceFolderId { get; set; } = "";
        public string Scope { get; set; } = "Metadados dos itens diretos da pasta selecionada. Sem corpos, assuntos ou bytes dos anexos; sem recursão ou alteração de itens. Não é fotografia transacional nem prova de completude no servidor.";
        public int ItemLimit { get; set; } = MaximumItems;
        public bool LimitReached { get; set; }
        public int? CountBefore { get; set; }
        public int? CountAfter { get; set; }
        public int UniqueEntryIds { get; set; }
        public int ItemsWithKnownSize { get; set; }
        public long ObservedSizeBytesSum { get; set; }
        public int ItemsWithAttachments { get; set; }
        public int ItemsWithUnknownAttachmentCount { get; set; }
        public long AttachmentCountObserved { get; set; }
        public Dictionary<int, int> AttachmentTypeCounts { get; set; } = new();
        public Dictionary<string, int> MessageClassCounts { get; set; } = new(StringComparer.Ordinal);
        public Dictionary<int, int> BodyFormatCounts { get; set; } = new();
        public List<CoverageItem> Items { get; set; } = new();
        public List<CoverageFailure> Failures { get; set; } = new();
    }
}
