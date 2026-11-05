using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

// A deliberately bounded mail-only pilot. Originals are never moved/deleted.
// Native Copy creates a distinct temporary item in the source folder; only
// that copy is moved. Failed/ambiguous operations are never retried automatically.
internal static class PilotCommand
{
    internal static int Sample(string[] args, bool useMsgTransport = false)
    {
        if (args.Length is < 5 or > 6)
        { Console.Error.WriteLine("Uso: pilot-sample <store-id> <folder-id> <novo.pst> <nova-execucao> <execucao-anterior> [limite]"); return 2; }
        return Execute(args.Take(4).Concat(new[] { args.Length == 6 ? args[5] : "10" }).ToArray(), true, args[4], useMsgTransport);
    }

    internal static int Execute(string[] args, bool collectBodyDifferences = false, string? previousRun = null, bool useMsgTransport = false, bool allowInitialDiagnosticCollection = false)
    {
        if (args.Length is < 4 or > 5 || (args.Length == 5 && !int.TryParse(args[4], out _)))
        { Console.Error.WriteLine("Uso: pilot-copy <store-id> <folder-entry-id> <destino.pst> <diretorio-execucao> [limite-1-a-25]"); return 2; }
        int limit = args.Length == 5 ? int.Parse(args[4]) : 5;
        object? app = null, session = null, source = null, items = null, stores = null;
        object? targetStore = null, root = null, folders = null, target = null;
        DurableJournal? journal = null;
        var report = new PilotReport { DiagnosticSampling = collectBodyDifferences,
            Transport = useMsgTransport ? "unicode-msg-direct-import" : "outlook-native-copy",
            VerificationVersion = useMsgTransport ? "msg-direct-import-v6-received-only" : "native-copy-msg-body-v3" };
        try
        {
            var pst = ValidateDestination(args[2], args[3], limit);
            var run = Path.GetFullPath(args[3]);
            var excluded = new HashSet<string>(StringComparer.Ordinal);
            if (collectBodyDifferences && previousRun is null && !allowInitialDiagnosticCollection)
                throw new ArgumentException("Amostragem requer execução anterior para reconciliação.");
            if (collectBodyDifferences && previousRun is not null)
            {
                var previous = JsonSerializer.Deserialize<PilotReport>(File.ReadAllText(Path.Combine(previousRun, "report.json")))
                    ?? throw new InvalidDataException("Relatório anterior inválido.");
                if (previous.SourceStoreId != args[0] || previous.SourceFolderId != args[1])
                    throw new InvalidOperationException("Execução anterior é de outra origem.");
                var audit = PilotAudit.Read(previousRun);
                if (audit.CopyCreationAmbiguous.Length != 0 || audit.ImportAmbiguous.Length != 0 ||
                    audit.TemporaryCopiesPossiblyRemaining.Length != 0 || audit.MetadataRestoreAmbiguous.Length != 0 || audit.Errors.Count != 0)
                    throw new InvalidOperationException("Execução anterior possui operações ambíguas; amostragem bloqueada.");
                foreach (var prior in previous.Items.Where(x => !string.IsNullOrEmpty(x.DestinationId))) excluded.Add(prior.SourceId);
                foreach (string prior in previous.ExcludedSourceIds) excluded.Add(prior);
                report.ExcludedSourceIds = excluded.ToArray();
                report.PreviousRun = Path.GetFullPath(previousRun);
            }
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            if ((bool)((dynamic)session).Offline) throw new InvalidOperationException("Outlook está offline.");
            source = ((dynamic)session).GetFolderFromID(args[1], args[0]);
            if (!string.Equals((string)((dynamic)source).StoreID, args[0], StringComparison.Ordinal))
                throw new InvalidOperationException("A pasta não pertence à caixa informada.");
            object? accessor = null;
            try
            {
                accessor = ((dynamic)source).PropertyAccessor;
                if ((int)((dynamic)accessor).GetProperty("http://schemas.microsoft.com/mapi/proptag/0x36010003") == 2)
                    throw new InvalidOperationException("Pasta de pesquisa não pode ser origem do piloto.");
            }
            finally { Release(accessor); }
            if ((int)((dynamic)source).DefaultItemType != 0)
                throw new InvalidOperationException("O piloto aceita somente pastas de e-mail.");
            items = ((dynamic)source).Items;
            report.SourceStoreId = args[0];
            report.SourceFolderId = args[1];
            report.SourceFolderPath = (string)((dynamic)source).FolderPath;
            report.SourceCountBefore = (int)((dynamic)items).Count;
            report.PstPath = pst;
            // Snapshot IDs before writing; later lookups never depend on mutable indices.
            var ids = new List<string>();
            long totalSize = 0;
            for (int i = 1; i <= report.SourceCountBefore && ids.Count < limit; i++)
            {
                object? item = null;
                try
                {
                    item = ((dynamic)items)[i];
                    if (excluded.Contains((string)((dynamic)item).EntryID)) continue;
                    if ((int)((dynamic)item).Class != 43 || (string)((dynamic)item).MessageClass != "IPM.Note")
                        throw new InvalidOperationException("Seleção contém item especial/protegido; piloto interrompido antes da criação do PST.");
                    if (!(bool)((dynamic)item).Sent)
                        throw new NotSupportedException("Mensagem não enviada detectada. O Outlook altera a data de recebimento na importação desse tipo; transferência bloqueada antes de criar o PST. O original será preservado.");
                    totalSize = checked(totalSize + (int)((dynamic)item).Size);
                    ids.Add((string)((dynamic)item).EntryID);
                }
                finally { Release(item); }
            }
            if (ids.Count == 0) { Console.WriteLine("Pasta vazia: nenhum PST criado."); return 0; }
            if (totalSize > 250L * 1024 * 1024) throw new InvalidOperationException("Piloto limitado a 250 MiB de itens estimados.");
            foreach (string volume in new[] { Path.GetPathRoot(pst)!, Path.GetPathRoot(run)! }.Distinct(StringComparer.OrdinalIgnoreCase))
                if (new DriveInfo(volume).AvailableFreeSpace < totalSize * 8 + 1024L * 1024 * 1024)
                    throw new IOException("Espaço livre insuficiente para o piloto e sua margem em " + volume);
            Directory.CreateDirectory(run);
            journal = new DurableJournal(Path.Combine(run, "journal.jsonl"));
            journal.Write(new { State = "planned", Report = report, ItemIds = ids, Limit = limit });
            // Recheck immediately before Outlook creates a file; never attach an existing PST.
            if (File.Exists(pst)) throw new IOException("O PST de destino passou a existir; operação cancelada.");
            journal.Write(new { State = "pst-create-started", Path = pst });
            ((dynamic)session).AddStoreEx(pst, 2); // olStoreUnicode
            stores = ((dynamic)session).Stores;
            for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
            {
                object? candidate = ((dynamic)stores)[i];
                try
                {
                    string filePath = (string)((dynamic)candidate).FilePath;
                    if (string.Equals(Path.GetFullPath(filePath), pst, StringComparison.OrdinalIgnoreCase))
                    { targetStore = candidate; candidate = null; break; }
                }
                catch (Exception ex) when (ex is ArgumentException or COMException)
                { journal.Write(new { State = "store-path-unavailable", Index = i, Error = ex.Message }); }
                finally { Release(candidate); }
            }
            if (targetStore is null) throw new InvalidOperationException("PST criado não foi localizado no perfil.");
            root = ((dynamic)targetStore).GetRootFolder();
            folders = ((dynamic)root).Folders;
            target = ((dynamic)folders).Add("Piloto - " + (string)((dynamic)source).Name, 6); // olFolderInbox
            report.TargetStoreId = (string)((dynamic)targetStore).StoreID;
            report.TargetFolderId = (string)((dynamic)target).EntryID;
            journal.Write(new { State = "destination-ready", report.TargetStoreId, report.TargetFolderId });
            int index = 0;
            foreach (string id in ids)
            {
                index++;
                var itemResult = new PilotItemResult { SourceId = id };
                report.Items.Add(itemResult);
                object? original = null, staged = null, imported = null, reopened = null;
                string itemDirectory = Path.Combine(run, $"item-{index:D3}");
                Directory.CreateDirectory(itemDirectory);
                try
                {
                    journal.Write(new { State = "item-started", Index = index, SourceId = id });
                    original = ((dynamic)session).GetItemFromID(id, args[0]);
                    object? parent = null;
                    try
                    {
                        parent = ((dynamic)original).Parent;
                        if ((string)((dynamic)parent).EntryID != args[1] || (string)((dynamic)parent).StoreID != args[0])
                            throw new InvalidOperationException("Item deixou a pasta selecionada; piloto interrompido.");
                    }
                    finally { Release(parent); }
                    string msg = Path.Combine(itemDirectory, "message.msg");
                    ((dynamic)original).SaveAs(msg, 9); // olMSGUnicode: source remains untouched.
                    string before = Snapshot(original, Path.Combine(itemDirectory, "source"), msg);
                    itemResult.SourceFingerprint = before;
                    if (useMsgTransport)
                    {
                        journal.Write(new { State = "local-msg-open-started", Index = index, SourceId = id,
                            MsgSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(msg))) });
                        staged = ((dynamic)session).OpenSharedItem(msg);
                        if ((int)((dynamic)staged).Class != 43 || (string)((dynamic)staged).MessageClass != "IPM.Note")
                            throw new InvalidOperationException("MSG intermediário não é uma mensagem comum.");
                    }
                    else
                    {
                    journal.Write(new { State = "source-copy-started", Index = index, SourceId = id, Fingerprint = before });
                    staged = ((dynamic)original).Copy();
                    itemResult.TemporarySourceCopyId = (string)((dynamic)staged).EntryID;
                    if (string.IsNullOrEmpty(itemResult.TemporarySourceCopyId) || itemResult.TemporarySourceCopyId == id)
                        throw new InvalidOperationException("A cópia não possui identidade distinta; movimentação bloqueada.");
                    journal.Write(new { State = "source-copy-created", Index = index, SourceId = id, itemResult.TemporarySourceCopyId });
                    object? copyParent = null;
                    try
                    {
                        copyParent = ((dynamic)staged).Parent;
                        if ((string)((dynamic)copyParent).EntryID != args[1] || (string)((dynamic)copyParent).StoreID != args[0])
                            throw new InvalidOperationException("A cópia temporária não está na pasta esperada; movimentação bloqueada.");
                    }
                    finally { Release(copyParent); }
                    }
                    journal.Write(new { State = "import-started", Index = index, SourceId = id, itemResult.TemporarySourceCopyId, Fingerprint = before });
                    imported = ((dynamic)staged).Move(target); // Only a new copy or local MSG object is moved.
                    itemResult.DestinationId = (string)((dynamic)imported).EntryID;
                    journal.Write(new { State = "imported", Index = index, itemResult.DestinationId });
                    reopened = ((dynamic)session).GetItemFromID(itemResult.DestinationId, report.TargetStoreId);
                    string destinationMsg = Path.Combine(itemDirectory, "destination-check.msg");
                    ((dynamic)reopened).SaveAs(destinationMsg, 9);
                    string destinationHash = Snapshot(reopened, Path.Combine(itemDirectory, "destination"), destinationMsg);
                    itemResult.BodyEvidence = SerializedBodyEvidence.Compare(msg, destinationMsg);
                    string afterMsg = Path.Combine(itemDirectory, "source-after.msg");
                    ((dynamic)original).SaveAs(afterMsg, 9);
                    string after = Snapshot(original, Path.Combine(itemDirectory, "source-after"), afterMsg);
                    if (before != after)
                        throw new InvalidDataException("Origem mudou nos campos conferidos; revisão necessária.");
                    if (before != destinationHash)
                    {
                        var differences = DifferentFields(Path.Combine(itemDirectory, "source", "field-hashes.json"),
                            Path.Combine(itemDirectory, "destination", "field-hashes.json"));
                        itemResult.DifferentFields = differences;
                        if (CanCollectBodyDifference(collectBodyDifferences, differences))
                        {
                            itemResult.Status = "copied-body-review-required";
                            itemResult.Error = "Cópia preservada para diagnóstico de corpo; não validada para remoção.";
                            journal.Write(new { State = itemResult.Status, Index = index, itemResult.SourceId, itemResult.DestinationId, Differences = differences });
                            Console.WriteLine($"Item {index}/{ids.Count}: copiado para diagnóstico; diferença de corpo registrada.");
                            continue;
                        }
                        throw new InvalidDataException("Destino ou origem mudou nos campos conferidos; revisão necessária.");
                    }
                    itemResult.Status = "verified-selected-fields";
                    journal.Write(new { State = itemResult.Status, Index = index, itemResult.SourceId, itemResult.DestinationId, Fingerprint = before });
                    Console.WriteLine($"Item {index}/{ids.Count}: cópia conferida nos campos selecionados.");
                }
                catch (Exception ex)
                {
                    itemResult.Status = "review-required";
                    itemResult.Error = ex.Message;
                    journal.Write(new { State = "review-required", Index = index, itemResult.SourceId, itemResult.DestinationId, Error = ex.Message });
                    // No automatic retry after an ambiguous import; avoids duplicate copies.
                    break;
                }
                finally { Release(reopened); Release(imported); Release(staged); Release(original); }
            }
            report.SourceCountAfter = (int)((dynamic)items).Count;
            report.FinishedUtc = DateTimeOffset.UtcNow;
            report.Status = !collectBodyDifferences && report.SourceCountBefore == report.SourceCountAfter && report.Items.Count == ids.Count && report.Items.All(x => x.Status == "verified-selected-fields")
                ? "pilot-completed" : "review-required";
            report.SampleCollectionCompleted = collectBodyDifferences && report.SourceCountBefore == report.SourceCountAfter &&
                report.Items.Count == ids.Count && report.Items.All(x => x.Status is "verified-selected-fields" or "copied-body-review-required");
            journal.Write(new { State = report.Status, report.SourceCountBefore, report.SourceCountAfter });
            WriteReport(run, report);
            Console.WriteLine($"{report.Status}: {report.Items.Count}/{ids.Count} itens processados. PST: {pst}\nOriginais preservados. Relatório: {Path.Combine(run, "report.json")}");
            return report.Status == "pilot-completed" ? 0 : 3;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Piloto interrompido: " + ex.Message);
            if (journal is not null)
            {
                report.Status = "review-required";
                report.Error = ex.Message;
                report.FinishedUtc = DateTimeOffset.UtcNow;
                try { journal.Write(new { State = "failed", Error = ex.Message }); }
                catch (Exception logError) { Console.Error.WriteLine("Falha ao persistir diário: " + logError.Message); }
                try { WriteReport(Path.GetFullPath(args[3]), report); }
                catch (Exception logError) { Console.Error.WriteLine("Falha ao persistir relatório: " + logError.Message); }
            }
            return 1;
        }
        finally
        {
            journal?.Dispose();
            Release(target); Release(folders); Release(root); Release(targetStore); Release(stores);
            Release(items); Release(source); Release(session); Release(app);
        }
    }

    private static string[] DifferentFields(string leftPath, string rightPath)
    {
        var left = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(leftPath))!;
        var right = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(rightPath))!;
        return left.Keys.Union(right.Keys).Where(key => left.GetValueOrDefault(key) != right.GetValueOrDefault(key)).Order(StringComparer.Ordinal).ToArray();
    }

    internal static bool CanCollectBodyDifference(bool diagnosticSampling, string[] differences) =>
        diagnosticSampling && differences.Length != 0 && differences.All(x => x.StartsWith("NativeBody-", StringComparison.Ordinal));


    internal static string ValidateDestination(string pstPath, string runPath, int limit)
    {
        if (limit is < 1 or > 25) throw new ArgumentOutOfRangeException(nameof(limit), "Limite permitido: 1 a 25.");
        if (!Path.IsPathFullyQualified(pstPath) || !Path.IsPathFullyQualified(runPath))
            throw new ArgumentException("Use caminhos absolutos para PST e execução.");
        string pst = Path.GetFullPath(pstPath), run = Path.GetFullPath(runPath);
        if (!string.Equals(Path.GetExtension(pst), ".pst", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("O destino precisa ter extensão .pst.");
        if (File.Exists(pst) || Directory.Exists(pst)) throw new IOException("Destino já existe; sobrescrita proibida.");
        if (File.Exists(run) || Directory.Exists(run)) throw new IOException("Diretório de execução já existe; reutilização bloqueada até reconciliação.");
        if (!Directory.Exists(Path.GetDirectoryName(pst)) || !Directory.Exists(Path.GetDirectoryName(run)))
            throw new DirectoryNotFoundException("As pastas pai dos destinos devem existir.");
        if (IsWithin(pst, run) || IsWithin(run, pst)) throw new ArgumentException("PST e diretório de execução devem ser separados.");
        foreach (string path in new[] { pst, run })
        {
            if (path.StartsWith(@"\\", StringComparison.Ordinal) || new DriveInfo(Path.GetPathRoot(path)!).DriveType != DriveType.Fixed)
                throw new ArgumentException("O piloto requer disco local fixo.");
            for (var parent = new DirectoryInfo(Path.GetDirectoryName(path)!); parent is not null; parent = parent.Parent)
            {
                if (parent.Attributes.HasFlag(FileAttributes.ReparsePoint)) throw new ArgumentException("Destino passa por link/redirecionamento de diretório.");
                if (parent.Name.Contains("OneDrive", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Destino não pode ficar no OneDrive.");
            }
            foreach (string variable in new[] { "OneDrive", "OneDriveCommercial", "OneDriveConsumer" })
            {
                var syncRoot = Environment.GetEnvironmentVariable(variable);
                if (!string.IsNullOrWhiteSpace(syncRoot) && IsWithin(path, syncRoot))
                    throw new ArgumentException("Destino não pode ficar no OneDrive.");
            }
        }
        return pst;
    }

    internal static bool IsWithin(string path, string directory)
    {
        string root = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string full = Path.GetFullPath(path);
        return full.StartsWith(root, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(full.TrimEnd(Path.DirectorySeparatorChar), root.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
    }

    internal static string Snapshot(object instance, string scratch, string msgPath)
    {
        dynamic item = instance;
        bool report = DeliveryReportEvidence.IsReport(instance);
        if (!report && ((int)item.Class != 43 || !DeliveryReportEvidence.SupportedMailClass((string)item.MessageClass)))
            throw new InvalidOperationException("Tipo de item não suportado no piloto.");
        Directory.CreateDirectory(scratch);
        var reportFields = report ? DeliveryReportEvidence.Read(msgPath) : null;
        var fields = new SortedDictionary<string, string>
        {
            ["MessageClass"] = (string)item.MessageClass,
            ["Subject"] = report ? reportFields!["Report-0037001F"] : (string)item.Subject ?? "",
            ["To"] = report ? "report-native" : (string)item.To ?? "", ["CC"] = report ? "report-native" : (string)item.CC ?? "", ["BCC"] = report ? "report-native" : (string)item.BCC ?? "",
            ["SenderEmailAddress"] = report ? "report-native" : (string)item.SenderEmailAddress ?? "",
            ["SentOn"] = report ? "report-native" : ((DateTime)item.SentOn).ToString("O"),
            ["ReceivedTime"] = report ? "report-native" : ((DateTime)item.ReceivedTime).ToString("O"),
            ["UnRead"] = ((bool)item.UnRead).ToString(),
            ["Importance"] = ((int)item.Importance).ToString(),
            ["Sensitivity"] = ((int)item.Sensitivity).ToString()
        };
        // Outlook may expose a sanitized/plain-text body in Junk Email. Compare
        // the serialized native body streams instead of context-dependent getters.
        if (report) foreach (var field in reportFields!) fields[field.Key] = field.Value;
        foreach (var field in MsgBodyFingerprint.Read(msgPath)) fields["NativeBody-" + field.Key] = field.Value;
        object? attachments = null;
        try
        {
            attachments = item.Attachments;
            int count = (int)((dynamic)attachments).Count;
            fields["AttachmentCount"] = count.ToString();
            for (int i = 1; i <= count; i++)
            {
                object? attachment = null;
                try
                {
                    attachment = ((dynamic)attachments)[i];
                    int attachmentType = (int)((dynamic)attachment).Type;
                    // Embedded messages can be exported intact. Keep exact byte
                    // equality; differing serializations stay blocked for review.
                    if (attachmentType is not (1 or 5))
                        throw new InvalidOperationException($"Anexo especial não suportado (tipo {attachmentType}); revisão necessária.");
                    if (attachmentType == 5) fields[$"Attachment-{i}-Type"] = "5";
                    string path = Path.Combine(scratch, $"attachment-{i:D3}." + (attachmentType == 5 ? "msg" : "bin"));
                    ((dynamic)attachment).SaveAsFile(path);
                    using var input = File.OpenRead(path);
                    fields[$"Attachment-{i}-Name"] = (string)((dynamic)attachment).FileName;
                    fields[$"Attachment-{i}-Hash"] = Convert.ToHexString(SHA256.HashData(input));
                    object? accessor = null;
                    try
                    {
                        accessor = ((dynamic)attachment).PropertyAccessor;
                        fields[$"Attachment-{i}-ContentId"] = OptionalProperty(accessor, "0x3712001F");
                        fields[$"Attachment-{i}-ContentLocation"] = OptionalProperty(accessor, "0x3713001F");
                        fields[$"Attachment-{i}-Hidden"] = OptionalProperty(accessor, "0x7FFE000B");
                    }
                    finally { Release(accessor); }
                }
                finally { Release(attachment); }
            }
        }
        finally { Release(attachments); }
        var fieldHashes = fields.ToDictionary(x => x.Key,
            x => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(x.Value))));
        File.WriteAllText(Path.Combine(scratch, "field-hashes.json"), JsonSerializer.Serialize(fieldHashes));
        return Fingerprint(fields);
    }

    private static string OptionalProperty(object accessor, string tag)
    {
        try
        {
            object value = ((dynamic)accessor).GetProperty("http://schemas.microsoft.com/mapi/proptag/" + tag);
            return "present:" + JsonSerializer.Serialize(value);
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8004010F)) { return "missing"; }
    }

    internal static int Review(string runDirectory)
    {
        object? app = null, session = null;
        try
        {
            string run = Path.GetFullPath(runDirectory);
            var report = JsonSerializer.Deserialize<PilotReport>(File.ReadAllText(Path.Combine(run, "report.json")))
                ?? throw new InvalidDataException("Relatório inválido.");
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            int index = 0;
            foreach (var result in report.Items)
            {
                index++;
                object? original = null, staged = null;
                string itemDir = Path.Combine(run, $"item-{index:D3}");
                string sourceDir = Path.Combine(itemDir, "review-source-" + Guid.NewGuid().ToString("N"));
                string stageDir = Path.Combine(itemDir, "review-stage-" + Guid.NewGuid().ToString("N"));
                try
                {
                    original = ((dynamic)session).GetItemFromID(result.SourceId, report.SourceStoreId);
                    staged = ((dynamic)session).OpenSharedItem(Path.Combine(itemDir, "message.msg"));
                    Directory.CreateDirectory(sourceDir);
                    Directory.CreateDirectory(stageDir);
                    string sourceMsg = Path.Combine(sourceDir, "source-review.msg");
                    string stagedMsg = Path.Combine(stageDir, "stage-review.msg");
                    ((dynamic)original).SaveAs(sourceMsg, 9);
                    ((dynamic)staged).SaveAs(stagedMsg, 9);
                    Snapshot(original, sourceDir, sourceMsg);
                    Snapshot(staged, stageDir, stagedMsg);
                    var sourceFields = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(sourceDir, "field-hashes.json")))!;
                    var stageFields = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(Path.Combine(stageDir, "field-hashes.json")))!;
                    var differences = sourceFields.Keys.Union(stageFields.Keys).Where(key => sourceFields.GetValueOrDefault(key) != stageFields.GetValueOrDefault(key)).ToArray();
                    Console.WriteLine($"Item {index}: campos divergentes entre origem atual e MSG: {string.Join(", ", differences)}");
                    Console.WriteLine($"Destino importado registrado: {!string.IsNullOrEmpty(result.DestinationId)}");
                    Console.WriteLine(JsonSerializer.Serialize(new
                    {
                        Body = BodyDiagnostics((string)((dynamic)original).Body, (string)((dynamic)staged).Body),
                        Html = BodyDiagnostics((string)((dynamic)original).HTMLBody, (string)((dynamic)staged).HTMLBody),
                        SourceFormat = (int)((dynamic)original).BodyFormat,
                        MsgFormat = (int)((dynamic)staged).BodyFormat
                    }));
                }
                finally { Release(staged); Release(original); }
            }
            object? destination = null, destinationItems = null, sourceFolder = null, sourceItems = null;
            try
            {
                if (!string.IsNullOrEmpty(report.TargetFolderId) && !string.IsNullOrEmpty(report.TargetStoreId))
                {
                    destination = ((dynamic)session).GetFolderFromID(report.TargetFolderId, report.TargetStoreId);
                    destinationItems = ((dynamic)destination).Items;
                    Console.WriteLine($"Itens atuais na pasta destino do piloto: {(int)((dynamic)destinationItems).Count}");
                }
                sourceFolder = ((dynamic)session).GetFolderFromID(report.SourceFolderId, report.SourceStoreId);
                sourceItems = ((dynamic)sourceFolder).Items;
                Console.WriteLine($"Itens atuais na origem: {(int)((dynamic)sourceItems).Count}");
            }
            finally { Release(sourceItems); Release(sourceFolder); Release(destinationItems); Release(destination); }
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine("Falha na revisão: " + ex.Message); return 1; }
        finally { Release(session); Release(app); }
    }

    private static object BodyDiagnostics(string? left, string? right)
    {
        left ??= ""; right ??= "";
        int prefix = 0;
        while (prefix < Math.Min(left.Length, right.Length) && left[prefix] == right[prefix]) prefix++;
        return new { SourceLength = left.Length, MsgLength = right.Length, CommonPrefix = prefix,
            EqualNormalizedLineEndings = left.Replace("\r\n", "\n") == right.Replace("\r\n", "\n"),
            EqualTrimmed = left.Trim() == right.Trim() };
    }

    internal static string Fingerprint(SortedDictionary<string, string> fields) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(fields))));

    private static void WriteReport(string run, PilotReport report)
    {
        string temporary = Path.Combine(run, "report-" + Guid.NewGuid().ToString("N") + ".partial.json");
        using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        { JsonSerializer.Serialize(stream, report, new JsonSerializerOptions { WriteIndented = true }); stream.Flush(true); }
        File.Move(temporary, Path.Combine(run, "report.json"), false);
    }

    internal static int SelfTest()
    {
        MsgBodyFingerprint.SelfTest();
        PilotAudit.SelfTest();
        MsgDifferenceAnalysis.SelfTest();
        PstReopenCheck.SelfTest();
        SerializedBodyEvidence.SelfTest();
        SyntheticReceivedMsg.SelfTest();
        if (CanCollectBodyDifference(false, ["NativeBody-HTML"]) ||
            CanCollectBodyDifference(true, []) || CanCollectBodyDifference(true, ["NativeBody-HTML", "Subject"]) ||
            CanCollectBodyDifference(true, ["Attachment-1-Hash"]) || !CanCollectBodyDifference(true, ["NativeBody-HTML"]))
            throw new Exception("Amostragem aceitou divergência fora do escopo diagnóstico.");
        string temp = Path.Combine(Path.GetTempPath(), "outlook-pilot-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        try
        {
            var a = new SortedDictionary<string, string> { ["Body"] = "Olá", ["AttachmentHash"] = "AAA" };
            var b = new SortedDictionary<string, string>(a);
            if (Fingerprint(a) != Fingerprint(b)) throw new Exception("Fingerprint instável.");
            b["AttachmentHash"] = "BBB";
            if (Fingerprint(a) == Fingerprint(b)) throw new Exception("Divergência de anexo ignorada.");
            b = new(a); b["Body"] = "Alterado";
            if (Fingerprint(a) == Fingerprint(b)) throw new Exception("Divergência de corpo ignorada.");
            string pst = Path.Combine(temp, "test.pst"), run = Path.Combine(temp, "run");
            ValidateDestination(pst, run, 5);
            MustReject(() => ValidateDestination(pst, run, 26));
            MustReject(() => ValidateDestination(Path.Combine(temp, "test.ost"), run, 5));
            MustReject(() => ValidateDestination("relative.pst", run, 5));
            File.WriteAllText(pst, "existing");
            MustReject(() => ValidateDestination(pst, run, 5));
            if (File.ReadAllText(pst) != "existing") throw new Exception("Arquivo sobrescrito.");
            Directory.CreateDirectory(run);
            MustReject(() => ValidateDestination(Path.Combine(temp, "other.pst"), run, 5));
            using (var journal = new DurableJournal(Path.Combine(temp, "test.jsonl")))
            { journal.Write(new { State = "planned" }); journal.Write(new { State = "import-started" }); }
            var lines = File.ReadAllLines(Path.Combine(temp, "test.jsonl"));
            if (lines.Length != 2 || JsonDocument.Parse(lines[1]).RootElement.GetProperty("Event").GetProperty("State").GetString() != "import-started")
                throw new Exception("Diário inválido.");
            Console.WriteLine("Piloto: testes de divergência, limites, caminhos, sobrescrita e diário aprovados. COM não exercitado.");
            return 0;
        }
        finally
        {
            // Only exact known test files; no recursive deletion.
            foreach (var file in new[] { "test.pst", "test.jsonl" }) File.Delete(Path.Combine(temp, file));
            if (Directory.Exists(Path.Combine(temp, "run"))) Directory.Delete(Path.Combine(temp, "run"));
            Directory.Delete(temp);
        }
    }
    private static void MustReject(Action action)
    {
        try { action(); } catch (ArgumentException) { return; } catch (IOException) { return; }
        throw new Exception("Validação aceitou configuração proibida.");
    }
    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);
}

internal sealed class DurableJournal : IDisposable
{
    private readonly FileStream stream;
    internal DurableJournal(string path) => stream = new(path, FileMode.CreateNew, FileAccess.Write, FileShare.Read);
    internal void Write<T>(T value)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new { Utc = DateTimeOffset.UtcNow, Event = value }) + "\n");
        stream.Write(bytes);
        stream.Flush(true);
    }
    public void Dispose() => stream.Dispose();
}
internal sealed class PilotReport
{
    public string Status { get; set; } = "planned";
    public string? Error { get; set; }
    public bool DiagnosticSampling { get; set; }
    public bool SampleCollectionCompleted { get; set; }
    public string? PreviousRun { get; set; }
    public string[] ExcludedSourceIds { get; set; } = [];
    public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset? FinishedUtc { get; set; }
    public string? SourceStoreId { get; set; }
    public string? SourceFolderId { get; set; }
    public string? SourceFolderPath { get; set; }
    public string? TargetStoreId { get; set; }
    public string? TargetFolderId { get; set; }
    public string? PstPath { get; set; }
    public int SourceCountBefore { get; set; }
    public int? SourceCountAfter { get; set; }
    public bool RemovalEnabled { get; set; } = false;
    public bool CompletenessVerified { get; set; } = false;
    public string Transport { get; set; } = "outlook-native-copy";
    public string VerificationVersion { get; set; } = "native-copy-msg-body-v3";
    public string VerificationScope { get; set; } = "Classe, assunto, blocos nativos do corpo nos MSGs serializados (presença/tamanho/SHA-256), destinatários textuais, remetente, datas, leitura, importância, sensibilidade, nomes e SHA-256 dos anexos comuns. Não comprova fidelidade da serialização inicial ao servidor nem conferência integral MAPI.";
    public List<PilotItemResult> Items { get; set; } = new();
}
internal sealed class PilotItemResult
{
    public string SourceId { get; set; } = "";
    public string? DestinationId { get; set; }
    public string? TemporarySourceCopyId { get; set; }
    public string? SourceFingerprint { get; set; }
    public string Status { get; set; } = "started";
    public string? Error { get; set; }
    public string[] DifferentFields { get; set; } = [];
    public SerializedBodyEvidence? BodyEvidence { get; set; }
}
