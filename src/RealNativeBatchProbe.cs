using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

// Explicit, bounded diagnostic. It never updates the real ledger, never removes
// originals and is deliberately separate from archive-run and the desktop UI.
internal static class RealNativeBatchProbe
{
    private const string OwnerProperty = "http://schemas.microsoft.com/mapi/string/{DDB565AD-3206-4E72-B044-C54157137867}/VolumeOwner";
    private const int RequiredItems = 5;
    internal static int Execute(string[] args)
    {
        if (args.Length != 4)
        {
            Console.Error.WriteLine("Uso: real-native-batch-probe <raiz> <execução> <novo-diretório-Downloads> <true-batch-100-report.json>");
            return 2;
        }
        string? run = null, error = null;
        object? app = null, session = null, sourceFolder = null, sourceStore = null;
        FileStream? rootLease = null;
        OutlookArchiveBackend? backend = null;
        bool createdRun = false;
        var selected = new List<ArchiveItem>();
        var original = new Dictionary<string, ArchiveSnapshot>(StringComparer.Ordinal);
        var baseline = new List<object>();
        var nativeTargets = new List<object>();
        var after = new List<object>();
        var all = Stopwatch.StartNew();
        double exportSeconds = 0, nativeProcessSeconds = 0, nativeValidationSeconds = 0, sourceRecheckSeconds = 0;
        bool nativeMatched = false, baselineMatched = false, sourceUnchanged = false;
        JsonElement? nativeResult = null;
        string? sourcePlanSha = null, prerequisiteSha = null;
        try
        {
            string root = Path.GetFullPath(args[0]).TrimEnd(Path.DirectorySeparatorChar);
            if (!Guid.TryParseExact(args[1], "N", out _)) throw new ArgumentException("Execução inválida.");
            string prerequisite = Path.GetFullPath(args[3]);
            ValidateSyntheticPrerequisite(prerequisite);
            prerequisiteSha = ArchiveFiles.Sha(prerequisite);
            string helper = Path.Combine(AppContext.BaseDirectory, "native-copy.exe");
            if (!File.Exists(helper)) throw new FileNotFoundException("Helper nativo ausente.", helper);
            run = Path.GetFullPath(args[2]);
            string downloads = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            ArchiveFiles.ValidateLocalDirectory(run);
            if (!PilotCommand.IsWithin(run, downloads) || run.TrimEnd(Path.DirectorySeparatorChar).Equals(downloads, StringComparison.OrdinalIgnoreCase) || run.Length > 120)
                throw new ArgumentException("Use uma pasta nova em Downloads, com caminho de até 120 caracteres.");
            if (File.Exists(run) || Directory.Exists(run)) throw new IOException("A pasta de diagnóstico já existe; não haverá repetição.");
            // Existing lock opened read-only: prevent a concurrent real operation,
            // without creating/updating the real catalog or its lock file.
            rootLease = new FileStream(Path.Combine(root, "archive.lock"), FileMode.Open, FileAccess.Read, FileShare.None);
            using var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite"), true);
            var job = ledger.GetJob(args[1]);
            if (!Path.GetFullPath(job.RootDirectory).Equals(Path.Combine(root, "jobs", job.Id), StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Execução fora da raiz informada.");
            var envelope = JsonSerializer.Deserialize<ArchivePlanEnvelope>(job.ConfigJson) ?? throw new InvalidDataException("Configuração ausente.");
            envelope.Config.Validate();
            if (!Path.GetFullPath(envelope.Config.RootDirectory).TrimEnd(Path.DirectorySeparatorChar).Equals(root, StringComparison.OrdinalIgnoreCase) || envelope.Config.SyntheticSourcePstPath is not null)
                throw new InvalidDataException("Este ensaio requer uma execução real existente nesta raiz.");
            string planPath = Path.Combine(job.RootDirectory, "source-plan.json");
            sourcePlanSha = ArchiveFiles.Sha(planPath);
            if (sourcePlanSha != envelope.PlanSha256) throw new InvalidDataException("O plano congelado foi alterado.");
            var frozen = (JsonSerializer.Deserialize<List<SourceDescriptor>>(File.ReadAllText(planPath)) ?? throw new InvalidDataException("Plano inválido."))
                .ToDictionary(x => ArchiveLedger.ComputeItemKey(x.StoreId, x.EntryId), StringComparer.Ordinal);
            var candidates = ledger.GetItems(job.Id).Where(x => x.State == ArchiveItemState.Planned && x.SourceStoreId == envelope.Config.SourceStoreId).Take(200).ToList();
            if (candidates.Count < RequiredItems) throw new InvalidDataException("Não há cinco itens planejados elegíveis para o ensaio.");
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            sourceFolder = ((dynamic)session).GetFolderFromID(candidates[0].SourceFolderId, envelope.Config.SourceStoreId);
            sourceStore = ((dynamic)sourceFolder).Store;
            if ((bool)((dynamic)sourceStore).IsDataFileStore || (int)((dynamic)sourceStore).ExchangeStoreType is not (0 or 1))
                throw new InvalidDataException("O ensaio exige uma pasta de caixa Exchange, não um PST local.");
            string folderId = (string)((dynamic)sourceFolder).EntryID;
            string storeId = (string)((dynamic)sourceFolder).StoreID;
            foreach (var row in candidates.Where(x => x.SourceFolderId == candidates[0].SourceFolderId))
            {
                if (!frozen.TryGetValue(row.Key, out var descriptor) || descriptor.StoreId != row.SourceStoreId || descriptor.FolderId != row.SourceFolderId ||
                    descriptor.EntryId != row.SourceEntryId || descriptor.RecordKey != row.SourceRecordKey || !descriptor.IsSupported || !descriptor.IsSent)
                    throw new InvalidDataException("Item fora das identidades suportadas do plano congelado.");
                if (!EligibleWithoutReceipts(session!, row, false)) continue;
                selected.Add(row);
                if (selected.Count == RequiredItems) break;
            }
            if (selected.Count != RequiredItems) throw new InvalidDataException("Não há cinco mensagens já lidas e sem solicitações ou notificações de recibo entre os primeiros 200 itens planejados da pasta.");
            if (selected.Any(x => !Same(session!, x.SourceStoreId, storeId) || !Same(session!, x.SourceFolderId, folderId)))
                throw new InvalidDataException("As identidades selecionadas não pertencem à mesma pasta.");
            Directory.CreateDirectory(run);
            createdRun = true;
            ArchiveFiles.WriteNew(Path.Combine(run, "scope.json"), new
            {
                JobId = job.Id, SourcePlanSha256 = sourcePlanSha, SyntheticPrerequisiteSha256 = prerequisiteSha,
                Count = selected.Count, CopyOnly = true, CatalogReadOnly = true, RemovalAllowed = false,
                ReceiptGuard = "Already read; message flags 0x300 clear; read/non-receipt requests false or absent",
                Items = selected.Select(x => new { x.Key, x.SourceStoreId, x.SourceFolderId, x.SourceEntryId, x.SourceRecordKey })
            });
            backend = new OutlookArchiveBackend();
            var clock = Stopwatch.StartNew();
            for (int i = 0; i < selected.Count; i++)
            {
                var snapshot = backend.Export(selected[i], Path.Combine(run, $"s{i:D2}"));
                original.Add(selected[i].Key, snapshot);
                selected[i].StageMsgPath = snapshot.MessagePath;
                selected[i].StageMsgSha256 = snapshot.MessageSha256;
            }
            exportSeconds = clock.Elapsed.TotalSeconds;
            if (original.Values.Sum(x => new FileInfo(x.MessagePath).Length) > 64L * 1024 * 1024)
                throw new InvalidDataException("O conjunto excede o limite de 64 MiB deste ensaio de cinco itens.");
            var nativeVolume = NewVolume(run, "native-only.pst", storeId);
            backend.CreateVolume(nativeVolume);
            var nativeFolder = CreateOwnedEmptyFolder(session!, nativeVolume, "Lote diagnóstico");
            try
            {
                // Recheck the no-receipt conditions immediately before the native
                // cross-store copy. Never call SetReadFlags on the originals.
                foreach (var row in selected)
                {
                    EligibleWithoutReceipts(session!, row, true);
                    if (original[row.Key].ChangeKey is not { Length: > 0 } key || !backend.SourceMatchesChangeKey(row, key))
                    {
                        var current = backend.SourceSnapshot(row, null, Path.Combine(run, "pre-" + row.Key[..10]));
                        if (current.Hash != original[row.Key].Hash) throw new InvalidDataException("Origem mudou antes da cópia nativa.");
                    }
                }
                string handoff = Path.Combine(run, "native-handoff.txt");
                File.WriteAllLines(handoff, new[] { "scoped-copy-only-v1", storeId, folderId, nativeVolume.StoreId!, nativeFolder.FolderId }
                    .Concat(selected.Select(x => x.SourceEntryId)), new UTF8Encoding(false));
                clock.Restart();
                using var process = new Process { StartInfo = new ProcessStartInfo(helper) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true } };
                process.StartInfo.ArgumentList.Add(handoff);
                process.Start();
                Task<string> stdoutTask = process.StandardOutput.ReadToEndAsync(), stderrTask = process.StandardError.ReadToEndAsync();
                process.WaitForExit();
                nativeProcessSeconds = clock.Elapsed.TotalSeconds;
                string stdout = stdoutTask.GetAwaiter().GetResult(), stderr = stderrTask.GetAwaiter().GetResult();
                File.WriteAllText(Path.Combine(run, "native-helper.json"), stdout);
                File.WriteAllText(Path.Combine(run, "native-helper-errors.txt"), stderr);
                try { nativeResult = JsonSerializer.Deserialize<JsonElement>(stdout); } catch (JsonException) { }
                // Even a failed call can have copied messages. Preserve every
                // result and stop; never retry or fall back after ambiguity.
                if (process.ExitCode != 0) throw new InvalidOperationException("Cópia nativa retornou falha; resultados preservados, sem repetição ou fallback.");
                if (nativeResult is not { } result || result.GetProperty("Requested").GetInt32() != RequiredItems ||
                    result.GetProperty("MoveFlag").GetBoolean() || result.GetProperty("HResult").GetString() != "0x0" || result.GetProperty("Stage").GetString() != "copy-returned")
                    throw new InvalidDataException("O helper não confirmou a cópia nativa exata de cinco itens.");
                clock.Restart();
                var captured = CaptureNative(session!, backend, nativeVolume, nativeFolder.FolderId, run, nativeTargets, original);
                nativeValidationSeconds = clock.Elapsed.TotalSeconds;
                nativeMatched = captured.OrderBy(x => x, StringComparer.Ordinal).SequenceEqual(original.Values.Select(x => x.Hash).OrderBy(x => x, StringComparer.Ordinal));
                if (!nativeMatched) throw new InvalidDataException("O multiconjunto dos campos estritos da cópia nativa difere da origem; baseline e originais preservados.");
            }
            finally { Release(nativeFolder.Folder); }
            // Paired MSG baseline: exactly the same immutable exported backups,
            // distinct new PST and the current backend's actual Import method.
            var baselineVolume = NewVolume(run, "msg-only.pst", storeId);
            backend.CreateVolume(baselineVolume);
            baselineMatched = true;
            for (int i = 0; i < selected.Count; i++)
            {
                var row = selected[i];
                row.RelativePath = ["Baseline diagnóstico"];
                row.Token = Guid.NewGuid().ToString("N"); // Local object only; no ledger write.
                clock.Restart();
                var identity = backend.Import(baselineVolume, row);
                double importSeconds = clock.Elapsed.TotalSeconds;
                clock.Restart();
                var target = backend.TargetSnapshot(baselineVolume, row, identity, Path.Combine(run, $"b{i:D2}"));
                double validationSeconds = clock.Elapsed.TotalSeconds;
                string[] differences = Differences(original[row.Key].Fields, target.Fields);
                baselineMatched &= differences.Length == 0;
                baseline.Add(new { row.Key, ImportSeconds = importSeconds, ValidationSeconds = validationSeconds, Matched = differences.Length == 0, Differences = differences });
            }
            clock.Restart();
            sourceUnchanged = true;
            for (int i = 0; i < selected.Count; i++)
            {
                var snapshot = backend.SourceSnapshot(selected[i], null, Path.Combine(run, $"a{i:D2}"));
                bool matched = snapshot.Hash == original[selected[i].Key].Hash;
                sourceUnchanged &= matched;
                after.Add(new { selected[i].Key, Unchanged = matched, Differences = Differences(original[selected[i].Key].Fields, snapshot.Fields) });
            }
            sourceRecheckSeconds = clock.Elapsed.TotalSeconds;
            if (!baselineMatched || !sourceUnchanged) throw new InvalidDataException("Conferência estrita da baseline ou da preservação dos originais não passou.");
            Console.WriteLine("Ensaio limitado concluído; cinco originais preservados, ambos os destinos conferidos. Nenhuma migração foi habilitada automaticamente.");
            return 0;
        }
        catch (Exception ex)
        {
            error = ex.GetType().Name + ": " + ex.Message;
            Console.Error.WriteLine(error);
            return 1;
        }
        finally
        {
            // Preservation evidence is collected even after an ambiguous helper
            // result or a body difference. A failed copy is never silently retried.
            if (backend is not null && run is not null && original.Count > 0 && after.Count == 0)
            {
                var preservation = Stopwatch.StartNew();
                sourceUnchanged = original.Count == selected.Count;
                for (int i = 0; i < selected.Count; i++)
                {
                    var row = selected[i];
                    if (!original.TryGetValue(row.Key, out var expected)) continue;
                    try
                    {
                        var snapshot = backend.SourceSnapshot(row, null, Path.Combine(run, $"f{i:D2}"));
                        bool matched = snapshot.Hash == expected.Hash;
                        sourceUnchanged &= matched;
                        after.Add(new { row.Key, Unchanged = matched, Differences = Differences(expected.Fields, snapshot.Fields) });
                    }
                    catch (Exception ex)
                    {
                        sourceUnchanged = false;
                        after.Add(new { row.Key, Unchanged = false, Error = ex.GetType().Name + ": " + ex.Message });
                    }
                }
                sourceRecheckSeconds = preservation.Elapsed.TotalSeconds;
            }
            backend?.Dispose();
            // Retain the private PSTs and evidence. Never detach profile stores or
            // clean up messages after a failed/partial cross-store operation.
            Release(sourceStore); Release(sourceFolder); Release(session); Release(app); rootLease?.Dispose();
            if (createdRun && run is not null && Directory.Exists(run) && File.Exists(Path.Combine(run, "scope.json")))
                ArchiveFiles.WriteNew(Path.Combine(run, "real-native-probe-report.json"), new
                {
                    CopyOnly = true, CatalogReadOnly = true, RemovalAllowed = false, ProductionEnabled = false,
                    Requested = selected.Count, SourcePlanSha256 = sourcePlanSha, SyntheticPrerequisiteSha256 = prerequisiteSha,
                    NativeMultisetMatched = nativeMatched, MsgBaselineMatched = baselineMatched, SourceUnchanged = sourceUnchanged,
                    ExportSeconds = exportSeconds, NativeProcessSeconds = nativeProcessSeconds, NativeHelper = nativeResult,
                    NativeTargetValidationSeconds = nativeValidationSeconds, SourceRecheckSeconds = sourceRecheckSeconds,
                    TotalSeconds = all.Elapsed.TotalSeconds, Error = error, NativeTargets = nativeTargets, MsgBaseline = baseline, SourceAfter = after
                });
        }
    }
    private static void ValidateSyntheticPrerequisite(string path)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var report = document.RootElement;
        if (report.GetProperty("SyntheticOnly").GetBoolean() != true || report.GetProperty("ProductionEnabled").GetBoolean() ||
            report.GetProperty("SourceCount").GetInt32() != 100 || report.GetProperty("TargetCount").GetInt32() != 100 ||
            !report.GetProperty("MultisetMatched").GetBoolean() || !report.GetProperty("SourceUnchanged").GetBoolean() ||
            !report.GetProperty("DuplicateMultiplicityPreserved").GetBoolean())
            throw new InvalidDataException("O ensaio artificial de cem mensagens precisa ter passado integralmente.");
        var helper = report.GetProperty("Transfer").GetProperty("Helper");
        if (helper.GetProperty("Requested").GetInt32() != 100 || helper.GetProperty("MoveFlag").GetBoolean() ||
            helper.GetProperty("Stage").GetString() != "copy-returned" || helper.GetProperty("HResult").GetString() != "0x0")
            throw new InvalidDataException("Transporte artificial sem confirmação de cópia nativa integral.");
    }
    private static ArchiveVolume NewVolume(string run, string filename, string sourceStore) => new()
    { JobId = Guid.NewGuid().ToString("N"), Number = 1, SourceStoreId = sourceStore, PstPath = Path.Combine(run, filename) };
    private static (object Folder, string FolderId) CreateOwnedEmptyFolder(object session, ArchiveVolume volume, string name)
    {
        object? store = null, root = null, folders = null, folder = null, accessor = null;
        try
        {
            store = ((dynamic)session).GetStoreFromID(volume.StoreId);
            if (!(bool)((dynamic)store).IsDataFileStore || !Path.GetFullPath((string)((dynamic)store).FilePath).Equals(volume.PstPath, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("PST privado fora do destino informado.");
            root = ((dynamic)store).GetRootFolder();
            accessor = ((dynamic)root).PropertyAccessor;
            if ((string)((dynamic)accessor).GetProperty(OwnerProperty) != volume.JobId + ":1") throw new InvalidDataException("PST sem proprietário deste ensaio.");
            Release(accessor); accessor = null;
            folders = ((dynamic)root).Folders;
            folder = ((dynamic)folders).Add(name, 6);
            accessor = ((dynamic)folder).PropertyAccessor;
            ((dynamic)accessor).SetProperty(OwnerProperty, volume.JobId + ":1");
            if ((string)((dynamic)accessor).GetProperty(OwnerProperty) != volume.JobId + ":1" || Count(folder!) != 0)
                throw new InvalidDataException("Pasta privada não está vazia ou identificada.");
            object result = folder; folder = null;
            return (result, (string)((dynamic)result).EntryID);
        }
        finally { Release(accessor); Release(folder); Release(folders); Release(root); Release(store); }
    }
    private static bool EligibleWithoutReceipts(object session, ArchiveItem row, bool required)
    {
        object? item = null, parent = null, accessor = null;
        try
        {
            item = ((dynamic)session).GetItemFromID(row.SourceEntryId, row.SourceStoreId);
            parent = ((dynamic)item).Parent;
            if (!Same(session, (string)((dynamic)parent).StoreID, row.SourceStoreId) ||
                !Same(session, (string)((dynamic)parent).EntryID, row.SourceFolderId) ||
                !Same(session, (string)((dynamic)item).EntryID, row.SourceEntryId) ||
                (int)((dynamic)item).Class != 43 || (string)((dynamic)item).MessageClass != "IPM.Note" || !(bool)((dynamic)item).Sent)
                throw new InvalidDataException("Identidade ou tipo do original não confere.");
            accessor = ((dynamic)item).PropertyAccessor;
            object flagsValue = ((dynamic)accessor).GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003");
            if (flagsValue is not int flags) throw new InvalidDataException("Sinalizadores da mensagem ausentes ou inválidos.");
            bool read = (flags & 1) != 0, pending = (flags & 0x300) != 0;
            bool requested = OptionalBool(accessor!, "0x0029000B") || OptionalBool(accessor!, "0x0C06000B");
            bool eligible = read && !pending && !requested;
            if (required && !eligible) throw new InvalidDataException("A condição de mensagem lida e sem recibos mudou; cópia bloqueada.");
            return eligible;
        }
        finally { Release(accessor); Release(parent); Release(item); }
    }
    private static bool OptionalBool(object accessor, string tag)
    {
        try
        {
            object value = ((dynamic)accessor).GetProperty("http://schemas.microsoft.com/mapi/proptag/" + tag);
            return value is bool result ? result : throw new InvalidDataException("Propriedade de recibo com tipo inválido.");
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8004010F)) { return false; }
    }
    private static List<string> CaptureNative(object session, OutlookArchiveBackend backend, ArchiveVolume volume, string folderId, string run, List<object> evidence, Dictionary<string, ArchiveSnapshot> originals)
    {
        object? folder = null, items = null;
        var hashes = new List<string>();
        try
        {
            folder = ((dynamic)session).GetFolderFromID(folderId, volume.StoreId);
            if ((string)((dynamic)folder).StoreID != volume.StoreId || Count(folder!) != RequiredItems)
                throw new InvalidDataException("O lote nativo não contém exatamente cinco cópias; nenhuma repetição.");
            items = ((dynamic)folder).Items;
            var identities = new List<DestinationIdentity>();
            for (int i = 1; i <= RequiredItems; i++)
            {
                object? item = null;
                try { item = ((dynamic)items)[i]; identities.Add(new(volume.StoreId!, folderId, (string)((dynamic)item).EntryID)); }
                finally { Release(item); }
            }
            if (identities.Select(x => x.EntryId).Distinct(StringComparer.Ordinal).Count() != RequiredItems) throw new InvalidDataException("IDs de destino repetidos.");
            for (int i = 0; i < identities.Count; i++)
            {
                var snapshot = backend.TargetSnapshot(volume, new ArchiveItem(), identities[i], Path.Combine(run, $"n{i:D2}"), false);
                hashes.Add(snapshot.Hash);
                evidence.Add(new
                {
                    identities[i].EntryId, Hash = snapshot.Hash,
                    FieldsPath = Path.Combine(Path.GetDirectoryName(snapshot.MessagePath)!, "field-hashes.json"),
                    // Diagnostic comparisons, not identity claims or dedup keys.
                    Comparisons = originals.Select(x => new { SourceKey = x.Key, ExpectedHash = x.Value.Hash, Matched = x.Value.Hash == snapshot.Hash, Differences = Differences(x.Value.Fields, snapshot.Fields) }).ToArray()
                });
            }
            if (Count(folder!) != RequiredItems) throw new InvalidDataException("O lote mudou durante a conferência.");
            return hashes;
        }
        finally { Release(items); Release(folder); }
    }
    private static string[] Differences(Dictionary<string, string> expected, Dictionary<string, string> actual) =>
        expected.Keys.Union(actual.Keys).Where(key => expected.GetValueOrDefault(key) != actual.GetValueOrDefault(key)).ToArray();
    private static bool Same(object session, string left, string right) => OutlookIdentity.AreSame(left, right, (a, b) => (bool)((dynamic)session).CompareEntryIDs(a, b));
    private static int Count(object folder) { object? items = null; try { items = ((dynamic)folder).Items; return (int)((dynamic)items).Count; } finally { Release(items); } }
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("oleaut32.dll", PreserveSig = true)] private static extern int GetActiveObject(ref Guid id, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? app);
}
