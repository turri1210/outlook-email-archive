using System.Runtime.InteropServices;

internal sealed partial class OutlookArchiveBackend : IArchiveBackend, IArchiveResumableScopeBackend
{
    private const string TokenProperty = "http://schemas.microsoft.com/mapi/string/{DDB565AD-3206-4E72-B044-C54157137867}/TransferToken";
    private const string OwnerProperty = "http://schemas.microsoft.com/mapi/string/{DDB565AD-3206-4E72-B044-C54157137867}/VolumeOwner";
    private const string RemovalTokenProperty = "http://schemas.microsoft.com/mapi/string/{DDB565AD-3206-4E72-B044-C54157137867}/SourceRemovalToken";
    internal bool RetainComparisonSnapshots { get; set; }
    private object? app, session;
    private readonly Dictionary<string, string> attached = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, TokenIndex> tokenIndexes = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> folderIds = new(StringComparer.Ordinal);
    private readonly Dictionary<string, FileStream> pstLeases = new(StringComparer.OrdinalIgnoreCase);
    private void LockPst(string path)
    {
        string full = Path.GetFullPath(path);
        if (pstLeases.ContainsKey(full)) return;
        string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookPstArchive", "locks");
        Directory.CreateDirectory(directory);
        string key = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(full.ToUpperInvariant())));
        pstLeases[full] = new FileStream(Path.Combine(directory, key + ".lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
    }
    private sealed class TokenIndex
    {
        internal Dictionary<string, List<DestinationIdentity>> Tokens = new(StringComparer.Ordinal);
        internal int Unmarked;
    }
    internal OutlookArchiveBackend()
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Outlook requer thread STA.");
        var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
        Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app)); session = ((dynamic)app!).Session;
    }
    internal object ListSources()
    {
        var result = new List<object>(); var warnings = new List<string>(); object? stores = null;
        try
        {
            stores = ((dynamic)session!).Stores;
            for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
            {
                object? store = null, root = null;
                try
                {
                    store = ((dynamic)stores)[i];
                    string path = (string)((dynamic)store).FilePath;
                    if (!string.IsNullOrWhiteSpace(path) && path.EndsWith(".pst", StringComparison.OrdinalIgnoreCase) && !File.Exists(path))
                    { warnings.Add("Arquivo de dados indisponível, ignorado: " + path); continue; }
                    root = ((dynamic)store).GetRootFolder();
                    var folders = new List<object>(); ListFolders(root, "", folders);
                    result.Add(new { Name = (string)((dynamic)store).DisplayName, StoreId = (string)((dynamic)store).StoreID, Folders = folders });
                }
                catch (COMException ex) { warnings.Add("Armazenamento indisponível, ignorado: " + ex.Message); }
                finally { Release(root); Release(store); }
            }
        }
        finally { Release(stores); }
        return new { Stores = result, Warnings = warnings };
    }
    private static void ListFolders(object folder, string parent, List<object> rows)
    {
        if (Property(folder, "http://schemas.microsoft.com/mapi/proptag/0x36010003") is int type && type == 2) return;
        string name = (string)((dynamic)folder).Name;
        string path = parent.Length == 0 ? name : parent + " > " + name;
        rows.Add(new { Name = name, FolderId = (string)((dynamic)folder).EntryID, Path = path });
        object? children = null;
        try
        {
            children = ((dynamic)folder).Folders;
            for (int i = 1; i <= (int)((dynamic)children).Count; i++)
            {
                object? child = null;
                try { child = ((dynamic)children)[i]; ListFolders(child, path, rows); }
                finally { Release(child); }
            }
        }
        finally { Release(children); }
    }
    public List<SourceDescriptor> Discover(ArchiveConfig config) => DiscoverFrozen(config, CaptureScope(config));
    public List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config)=>CaptureScope(config,[],_=>{});
    public List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config,IReadOnlyList<FolderScopeSnapshot> saved,Action<FolderScopeSnapshot> checkpoint)
    {
        object? folder = null; var scope = new List<FolderScopeSnapshot>();
        try
        {
            folder = ((dynamic)session!).GetFolderFromID(config.SourceFolderId, config.SourceStoreId);
            if ((string)((dynamic)folder).StoreID != config.SourceStoreId) throw new ArchiveReviewException("Pasta de origem inválida.");
            if (config.SyntheticSourcePstPath is not null)
            {
                object? store = null;
                try
                {
                    store = ((dynamic)folder).Store;
                    if (!string.Equals(Path.GetFullPath((string)((dynamic)store).FilePath), Path.GetFullPath(config.SyntheticSourcePstPath), StringComparison.OrdinalIgnoreCase) ||
                        Path.GetExtension(config.SyntheticSourcePstPath) != ".pst") throw new ArchiveReviewException("Origem artificial não é o PST informado.");
                }
                finally { Release(store); }
            }
            if (config.PreserveFolderStructure) { InitializeFolderStructure(config.SourceStoreId); config.FolderMappings.Clear(); }
            CaptureFolders(folder, [], config, scope,saved.ToDictionary(x=>x.FolderId),checkpoint);
            if(saved.Any(x=>!scope.Any(y=>y.FolderId==x.FolderId)))throw new ArchiveReviewException("Uma pasta já capturada não está mais na origem. Escopo preservado para revisão.");
            return scope;
        }
        finally { Release(folder); }
    }
    private void CaptureFolders(object folder, string[] parent, ArchiveConfig config, List<FolderScopeSnapshot> scope,IReadOnlyDictionary<string,FolderScopeSnapshot> saved,Action<FolderScopeSnapshot> checkpoint)
    {
        object? children = null;
        try
        {
            if (Property(folder, "http://schemas.microsoft.com/mapi/proptag/0x36010003") is int type && type == 2) return;
            string[] path = [.. parent, (string)((dynamic)folder).Name];
            if (config.PreserveFolderStructure) RegisterFolderDestination(folder, path, config);
            string folderId=(string)((dynamic)folder).EntryID;
            bool retained=saved.TryGetValue(folderId,out var previous);
            if(retained && !previous!.RelativePath.SequenceEqual(path))throw new ArchiveReviewException("Uma pasta já capturada foi renomeada ou movida. Escopo preservado para revisão.");
            var ids = retained ? previous!.EntryIds.ToList() : CaptureEntryIds(folder, config.MaxItems);
            if ((long)scope.Sum(x => x.EntryIds.Count) + ids.Count > config.MaxItems)
                throw new ArchiveReviewException("Limite de identidades do escopo atingido; nenhum plano parcial será criado.");
            var snapshot=new FolderScopeSnapshot(folderId,path,ids);scope.Add(snapshot);if(!retained)checkpoint(snapshot);
            Console.WriteLine($"Planejamento: escopo inicial de {ids.Count:N0} itens capturado — {string.Join(" > ", path)}");
            if (!config.IncludeSubfolders) return;
            children = ((dynamic)folder).Folders; int count = (int)((dynamic)children).Count;
            for (int i = 1; i <= count; i++)
            {
                object? child = null;
                try { child = ((dynamic)children)[i]; CaptureFolders(child, path, config, scope,saved,checkpoint); }
                finally { Release(child); }
            }
        }
        finally { Release(children); }
    }
    private bool SameIdentity(string? left, string? right) => OutlookIdentity.AreSame(left, right,
        (a, b) => (bool)((dynamic)session!).CompareEntryIDs(a, b));
    private List<string> CaptureEntryIds(object folder, int limit)
    {
        object? table = null, columns = null, column = null, longColumn = null;
        try
        {
            table = ((dynamic)folder).GetTable("", 0); // olUserItems: ordinary, visible items only.
            columns = ((dynamic)table).Columns; ((dynamic)columns).RemoveAll();
            column = ((dynamic)columns).Add("EntryID");
            try { longColumn = ((dynamic)columns).Add("http://schemas.microsoft.com/mapi/proptag/0x66700102"); }
            catch (COMException) { /* Provider may not expose PR_LONGTERM_ENTRYID_FROM_TABLE. Resolve the captured IDs instead. */ }
            int count = (int)((dynamic)table).GetRowCount();
            if (count < 0 || count > limit) throw new ArchiveReviewException("Limite de identidades da pasta atingido.");
            if (count == 0) return [];
            // One bulk read, then stop consulting the live table. Never enumerate
            // Items[i] while arrivals can shift its positions over several minutes.
            object data = ((dynamic)table).GetArray(count);
            if (data is not Array array || array.Rank != 2 || array.GetLength(0) != count)
                throw new ArchiveReviewException("Escopo inicial incompleto; algum item original pode ter sido removido durante a captura.");
            var ids = new List<string>(count);
            for (int i = array.GetLowerBound(0); i <= array.GetUpperBound(0); i++)
            {
                if (array.GetValue(i, array.GetLowerBound(1)) is not string id || string.IsNullOrWhiteSpace(id))
                    throw new ArchiveReviewException("Identidade inválida na captura do escopo inicial.");
                string? persistent = null;
                if (longColumn is not null && array.GetLength(1) > 1)
                {
                    object? value = array.GetValue(i, array.GetLowerBound(1) + 1);
                    if (value is byte[] bytes && bytes.Length > 0) persistent = Convert.ToHexString(bytes);
                    else if (value is string text && !string.IsNullOrWhiteSpace(text)) persistent = text;
                }
                if (persistent is null)
                {
                    object? item = null;
                    try
                    {
                        item = ((dynamic)session!).GetItemFromID(id, (string)((dynamic)folder).StoreID);
                        persistent = (string)((dynamic)item).EntryID;
                        if (!SameIdentity(id, persistent)) throw new ArchiveReviewException("Não foi possível confirmar a identidade persistente de um item do escopo.");
                    }
                    finally { Release(item); }
                }
                ids.Add(persistent.ToUpperInvariant());
            }
            if (ids.Distinct(StringComparer.Ordinal).Count() != ids.Count) throw new ArchiveReviewException("Identidades duplicadas no escopo inicial.");
            return ids;
        }
        finally { Release(longColumn); Release(column); Release(columns); Release(table); }
    }
    public List<SourceDescriptor> DiscoverFrozen(ArchiveConfig config, IReadOnlyList<FolderScopeSnapshot> scope)
    {
        var rows = new List<SourceDescriptor>();
        foreach (var folder in scope)
        {
            int count = folder.EntryIds.Count, scanned = 0;
            if(count>1)Console.WriteLine($"Planejamento: 0/{count:N0} itens do escopo inicial — {string.Join(" > ", folder.RelativePath)}");
            for (int index = 0; index < folder.EntryIds.Count; index++)
            {
                string id = folder.EntryIds[index];
                object? item = null, parent = null;
                try
                {
                    item = ((dynamic)session!).GetItemFromID(id, config.SourceStoreId);
                    parent = ((dynamic)item).Parent;
                    string persistentId = ((string)((dynamic)item).EntryID).ToUpperInvariant();
                    if (!SameIdentity((string)((dynamic)parent).StoreID, config.SourceStoreId))
                        throw new ArchiveReviewException("Item do escopo inicial pertence a outro armazenamento; processamento bloqueado.");
                    if (!SameIdentity((string)((dynamic)parent).EntryID, folder.FolderId))
                        throw new ArchiveReviewException("Item do escopo inicial foi movido para outra pasta; processamento bloqueado.");
                    if (!SameIdentity(persistentId, id))
                        throw new ArchiveReviewException("O Outlook não confirmou que a identidade capturada corresponde ao item aberto; processamento bloqueado.");
                    folder.EntryIds[index] = persistentId;
                    string messageClass = (string)((dynamic)item).MessageClass;
                    bool report = DeliveryReportEvidence.IsReport(item);
                    bool supported = report || (int)((dynamic)item).Class == 43 && DeliveryReportEvidence.SupportedMailClass(messageClass);
                    bool sent = report || supported && (bool)((dynamic)item).Sent;
                    DateTime? messageDateUtc = null; string? dateBasis = null;
                    if (config.CutoffDateInclusive is not null && supported && sent)
                    {
                        DateTime? ReadDate(string tag) => Property(item, "http://schemas.microsoft.com/mapi/proptag/" + tag) is DateTime value
                            ? DateTime.SpecifyKind(value, DateTimeKind.Utc) : null;
                        (messageDateUtc, dateBasis) = ArchiveDateFilter.SelectDate(ReadDate("0x0E060040"), ReadDate("0x00390040"));
                    }
                    if (supported && sent && !ArchiveDateFilter.Includes(messageDateUtc, config.CutoffDateInclusive, config.CutoffTimeZoneId)) continue;
                    if (rows.Count >= config.MaxItems) throw new ArchiveReviewException("Limite de descoberta atingido; nenhum plano parcial será criado.");
                    rows.Add(new(config.SourceStoreId, folder.FolderId, persistentId, folder.RelativePath,
                        (int)((dynamic)item).Size, supported, sent, BinaryProperty(item, "0x0FF90102"), messageDateUtc, dateBasis, messageClass));
                }
                finally
                {
                    Release(parent); Release(item); scanned++;
                    if (count>1 && (scanned % 100 == 0 || scanned == count))
                        Console.WriteLine($"Planejamento: {scanned:N0}/{count:N0} do escopo inicial analisados; {rows.Count:N0} selecionados — {string.Join(" > ", folder.RelativePath)}");
                }
            }
        }
        return rows;
    }
    public ArchiveSnapshot Export(ArchiveItem item, string evidenceDirectory)
    {
        using var timing=ArchiveOperationTiming.Measure("source.export");
        object? source = null;
        try { source = OpenSource(item, item.SourceEntryId, item.SourceFolderId, item.SourceRecordKey); return Snapshot(source, evidenceDirectory, true, item.RepairToken is not null); }
        finally { Release(source); }
    }
    public ArchiveSnapshot SourceSnapshot(ArchiveItem item, SourceLocation? location, string evidenceDirectory)
    {
        object? source = null;
        try
        {
            source = OpenSource(item, location?.EntryId ?? item.SourceEntryId, location?.FolderId ?? item.SourceFolderId,
                location?.RecordKey ?? item.SourceRecordKey);
            return Snapshot(source, evidenceDirectory, false, item.RepairToken is not null);
        }
        finally { Release(source); }
    }
    public bool SourceMatchesChangeKey(ArchiveItem item, string expectedChangeKey)
    {
        object? source = null;
        try
        {
            if (item.SourceMessageClass is DeliveryReportEvidence.MessageClass or DeliveryReportEvidence.AutomaticReplyClass)
            { source = OpenSource(item,item.SourceEntryId,item.SourceFolderId,item.SourceRecordKey); return !string.IsNullOrEmpty(expectedChangeKey) && BinaryProperty(source,"0x65E20102") == expectedChangeKey; }
            source = OpenFastSource(item, [], out var values);
            return !string.IsNullOrEmpty(expectedChangeKey) && values[5] is byte[] changeKey && Convert.ToHexString(changeKey) == expectedChangeKey;
        }
        finally { Release(source); }
    }
    public ArchiveSnapshot BackupSnapshot(ArchiveItem item, string evidenceDirectory)
    {
        using var timing=ArchiveOperationTiming.Measure("backup.read");
        if(ArchiveEvidenceLease.TryBackup(item,out var retained))return retained;
        object? source = null;
        string working = Path.Combine(evidenceDirectory, "opened.msg");
        string before = ArchiveFiles.Sha(item.StageMsgPath!);
        try
        {
            Directory.CreateDirectory(evidenceDirectory);
            // Outlook retains native MSG handles after releasing the COM wrapper.
            // Open only a disposable copy; parse a separate, unopened copy so the
            // compound-file reader never competes with Outlook for a file handle.
            bool guarded=ArchiveEvidenceLease.Guards(item.StageMsgPath!);
            string serialized = guarded ? item.StageMsgPath! : Path.Combine(evidenceDirectory, "serialized.msg");
            if(!guarded)File.Copy(item.StageMsgPath!, serialized, false);
            if (ArchiveFiles.Sha(serialized) != before) throw new ArchiveReviewException("Backup mudou durante a leitura.");
            File.Copy(serialized, working, false);
            if(ArchiveFiles.Sha(working)!=before)throw new ArchiveReviewException("Cópia de trabalho do backup não confere.");
            using(ArchiveOperationTiming.Measure("backup.open-shared-msg"))source = ((dynamic)session!).OpenSharedItem(working);
            using(ArchiveOperationTiming.Measure("backup.fields-attachments"))PilotCommand.Snapshot(source, evidenceDirectory, serialized);
            var fields = ArchiveFiles.ReadFields(Path.Combine(evidenceDirectory, "field-hashes.json"));
            var relevant = RelevantMessageValidation.Generate(source, serialized, evidenceDirectory, recoveryBodyContext: item.RepairToken is not null);
            CleanupAttachments(evidenceDirectory);
            string after = ArchiveFiles.Sha(item.StageMsgPath!);
            if (before != after) throw new ArchiveReviewException("Backup MSG mudou durante a conferência; remoção bloqueada.");
            if(!guarded)File.Delete(serialized);
            var snapshot=new ArchiveSnapshot(fields,item.StageMsgPath!,after,null,relevant);
            ArchiveEvidenceLease.RememberBackup(item,snapshot);return snapshot;
        }
        finally
        {
            if (source is not null) { try { ((dynamic)source).Close(1); } catch (COMException) { } }
            Release(source);
            try { if (File.Exists(working)) File.Delete(working); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }
    private static ArchiveSnapshot Snapshot(object item, string directory, bool keepMessage, bool recoveryBodyContext = false)
    {
        using var timing=ArchiveOperationTiming.Measure("snapshot.total");
        if (Directory.Exists(directory)) throw new IOException("Diretório de evidência já existe.");
        Directory.CreateDirectory(directory);
        string? before = BinaryProperty(item, "0x65E20102");
        string msg = Path.Combine(directory, "message.msg");
        using(ArchiveOperationTiming.Measure("snapshot.save-msg"))((dynamic)item).SaveAs(msg, 9);
        using (var file = new FileStream(msg, FileMode.Open, FileAccess.ReadWrite, FileShare.Read)) file.Flush(true);
        using(ArchiveOperationTiming.Measure("snapshot.fields-attachments"))PilotCommand.Snapshot(item, directory, msg);
        Dictionary<string,string>? relevant;
        using(ArchiveOperationTiming.Measure("snapshot.relevant-profile"))relevant = RelevantMessageValidation.Generate(item, msg, directory, recoveryBodyContext: recoveryBodyContext);
        string? after = BinaryProperty(item, "0x65E20102");
        if (before != after) throw new ArchiveReviewException("ChangeKey da mensagem mudou durante a conferência.");
        var result = new ArchiveSnapshot(ArchiveFiles.ReadFields(Path.Combine(directory, "field-hashes.json")), msg, ArchiveFiles.Sha(msg), after, relevant);
        CleanupAttachments(directory);
        if (!keepMessage)
        {
            try { File.Delete(msg); } // Only this fresh diagnostic copy; never the retained backup.
            catch (IOException) { /* Outlook/antivirus can retain a temporary handle; keep evidence instead of failing validation. */ }
            catch (UnauthorizedAccessException) { }
        }
        return result;
    }
    private static void CleanupAttachments(string directory)
    {
        // PilotSnapshot generates exactly these numbered scratch files. Retain
        // hashes, not additional full attachment copies for a 50 GB operation.
        foreach (string file in Directory.EnumerateFiles(directory, "attachment-*.bin", SearchOption.TopDirectoryOnly)) File.Delete(file);
    }
    private void CheckDestinationStore(object store, ArchiveVolume volume)
    {
        if ((string)((dynamic)store).StoreID == volume.SourceStoreId)
            throw new ArchiveReviewException("O PST de destino não pode ser o armazenamento da própria origem.");
        if (!(bool)((dynamic)store).IsDataFileStore || !string.Equals(Path.GetFullPath((string)((dynamic)store).FilePath), Path.GetFullPath(volume.PstPath), StringComparison.OrdinalIgnoreCase))
            throw new ArchiveReviewException("O armazenamento não corresponde ao PST selecionado.");
        object? defaultStore = null;
        try
        {
            defaultStore = ((dynamic)session!).DefaultStore;
            if ((string)((dynamic)defaultStore).StoreID == (string)((dynamic)store).StoreID)
                throw new ArchiveReviewException("O arquivo de dados padrão do perfil não pode ser usado como destino nesta versão.");
        }
        finally { Release(defaultStore); }
    }
    private object OwnedScope(object store, ArchiveVolume volume, bool create)
    {
        object root = ((dynamic)store).GetRootFolder();
        if (!volume.IsExisting && !volume.PreserveFolderStructure) return root;
        object? folders = null, found = null;
        try
        {
            folders = ((dynamic)root).Folders;
            string name = volume.PreserveFolderStructure ? "Controle de arquivamento-" + volume.JobId : "Arquivo-" + volume.JobId;
            if (volume.PreserveFolderStructure && volume.ControlFolderId is not null)
            {
                found = ((dynamic)session!).GetFolderFromID(volume.ControlFolderId, volume.StoreId);
                if (!SameIdentity((string)((dynamic)found).StoreID, volume.StoreId) || (string?)Property(found, OwnerProperty) != Owner(volume)) throw new ArchiveReviewException("Controle de arquivamento não corresponde à execução.");
                object known = found; found = null; return known;
            }
            for (int i = 1; i <= (int)((dynamic)folders).Count; i++)
            {
                object candidate = ((dynamic)folders)[i];
                if ((string)((dynamic)candidate).Name == name)
                {
                    if (found is not null) { Release(candidate); throw new ArchiveReviewException("Pasta de arquivamento ambígua no PST existente."); }
                    found = candidate;
                }
                else Release(candidate);
            }
            if (found is null && create)
            {
                found = ((dynamic)folders).Add(name, 6);
                SetProperty(found, OwnerProperty, Owner(volume));
            }
            if (found is null || (string?)Property(found, OwnerProperty) != Owner(volume))
                throw new ArchiveReviewException("Pasta própria do arquivamento ausente ou sem identificação válida no PST existente.");
            if (volume.PreserveFolderStructure) { volume.ControlFolderId = (string)((dynamic)found).EntryID; SetBatchHidden(found, true); }
            object result = found; found = null; return result;
        }
        finally { Release(found); Release(folders); Release(root); }
    }
    private void EnrollExisting(ArchiveVolume volume)
    {
        ArchiveFiles.ValidateExistingPst(volume.PstPath);
        object? store = null, owned = null;
        try
        {
            using var guard = new FileStream(volume.PstPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            store = FindStore(volume.PstPath, volume.StoreId);
            if (store is null) { ((dynamic)session!).AddStoreEx(volume.PstPath, 2); store = FindStore(volume.PstPath, volume.StoreId); }
            if (store is null) throw new ArchiveReviewException("PST selecionado não foi localizado no Outlook.");
            CheckDestinationStore(store, volume);
            owned = OwnedScope(store, volume, true);
            volume.StoreId = (string)((dynamic)store).StoreID; attached[volume.PstPath] = volume.StoreId;
        }
        finally { Release(owned); Release(store); }
    }
    public void CreateVolume(ArchiveVolume volume)
    {
        LockPst(volume.PstPath);
        if (volume.IsExisting) { EnrollExisting(volume); return; }
        if (File.Exists(volume.PstPath)) { AttachVolume(volume); return; }
        object? store = null, root = null;
        try
        {
            ((dynamic)session!).AddStoreEx(volume.PstPath, 2);
            store = FindStore(volume.PstPath, volume.StoreId) ?? throw new ArchiveReviewException("PST criado não localizado.");
            root = volume.PreserveFolderStructure ? OwnedScope(store, volume, true) : ((dynamic)store).GetRootFolder();
            SetProperty(root, OwnerProperty, Owner(volume)); // Folder properties persist without MailItem.Save.
            if ((string?)Property(root, OwnerProperty) != Owner(volume)) throw new ArchiveReviewException("PST sem proprietário confirmado.");
            volume.StoreId = (string)((dynamic)store).StoreID;
            attached[volume.PstPath] = volume.StoreId;
        }
        finally { Release(root); Release(store); }
    }
    public void AttachVolume(ArchiveVolume volume)
    {
        using var timing=ArchiveOperationTiming.Measure("pst.attach-and-check");
        LockPst(volume.PstPath);
        if (!File.Exists(volume.PstPath)) throw new FileNotFoundException("PST registrado está ausente; substituição automática bloqueada.");
        if (attached.TryGetValue(volume.PstPath, out var cached)) { volume.StoreId = cached; return; }
        object? store = null, root = null;
        try
        {
            using var guard = new FileStream(volume.PstPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            store = FindStore(volume.PstPath, volume.StoreId);
            if (store is null) { ((dynamic)session!).AddStoreEx(volume.PstPath, 2); store = FindStore(volume.PstPath, volume.StoreId); }
            if (store is null) throw new ArchiveReviewException("PST não localizado.");
            CheckDestinationStore(store, volume);
            root = OwnedScope(store, volume, false);
            if ((string?)Property(root, OwnerProperty) != Owner(volume)) throw new ArchiveReviewException("PST existente não possui o marcador desta execução; reutilização bloqueada.");
            volume.StoreId = (string)((dynamic)store).StoreID;
            attached[volume.PstPath] = volume.StoreId;
        }
        finally { Release(root); Release(store); }
    }
    public long MeasureVolume(ArchiveVolume volume) => new FileInfo(volume.PstPath).Length;
    public TokenLookup FindToken(ArchiveVolume volume, ArchiveItem item)
    {
        using var timing=ArchiveOperationTiming.Measure("target.find-token");
        AttachVolume(volume);
        object? folder = null, store = null, root = null;
        try
        {
            folder = TargetFolder(volume, item.RelativePath, false);
            if (!tokenIndexes.TryGetValue(volume.PstPath, out var index))
            {
                index = new TokenIndex(); store = FindStore(volume.PstPath, volume.StoreId)!;
                root = volume.PreserveFolderStructure ? ((dynamic)store).GetRootFolder() : OwnedScope(store, volume, false);
                ScanTokens(root, volume.StoreId!, index); tokenIndexes[volume.PstPath] = index;
            }
            var matches = index.Tokens.GetValueOrDefault(item.DestinationToken) ?? [];
            if (matches.Count > 0 && (folder is null || matches.Any(x => x.FolderId != (string)((dynamic)folder).EntryID)))
                throw new ArchiveReviewException("Cópia identificada foi movida externamente para outra pasta do PST.");
            return new(matches.ToList(), index.Unmarked);
        }
        finally { Release(root); Release(store); Release(folder); }
    }
    private static void ScanTokens(object folder, string storeId, TokenIndex index)
    {
        object? table = null, columns = null, children = null;
        try
        {
            if (Property(folder, "http://schemas.microsoft.com/mapi/proptag/0x36010003") is int type && type == 2) return;
            string id = (string)((dynamic)folder).EntryID;
            table = ((dynamic)folder).GetTable("", 0); columns = ((dynamic)table).Columns;
            ((dynamic)columns).RemoveAll(); ((dynamic)columns).Add("EntryID"); ((dynamic)columns).Add(TokenProperty + "/0x0000001f");
            int count = (int)((dynamic)table).GetRowCount(), scanned = 0;
            while (!(bool)((dynamic)table).EndOfTable)
            {
                Array rows = (Array)((dynamic)table).GetArray(250);
                for (int i = rows.GetLowerBound(0); i <= rows.GetUpperBound(0); i++)
                {
                    string? entryId = rows.GetValue(i, rows.GetLowerBound(1)) as string;
                    string? token = rows.GetValue(i, rows.GetLowerBound(1) + 1) as string;
                    if (string.IsNullOrWhiteSpace(entryId)) throw new ArchiveReviewException("Identificador ausente no inventário do PST.");
                    if (string.IsNullOrEmpty(token)) index.Unmarked++;
                    else
                    {
                        if (!index.Tokens.TryGetValue(token, out var matches)) index.Tokens[token] = matches = [];
                        matches.Add(new(storeId, id, entryId));
                    }
                    scanned++;
                }
                Console.WriteLine($"Inventário de identificadores no PST: {scanned:N0}/{count:N0} mensagens nesta pasta…");
            }
            if (scanned != count || (int)((dynamic)table).GetRowCount() != count) throw new ArchiveReviewException("PST mudou durante o inventário de tokens.");
            children = ((dynamic)folder).Folders;
            for (int i = 1; i <= (int)((dynamic)children).Count; i++)
            {
                object? child = null;
                try { child = ((dynamic)children)[i]; ScanTokens(child, storeId, index); }
                finally { Release(child); }
            }
        }
        finally { Release(children); Release(columns); Release(table); }
    }
    public DestinationIdentity Import(ArchiveVolume volume, ArchiveItem item)
    {
        using var timing=ArchiveOperationTiming.Measure("target.import");
        if(item.SourceMessageClass==DeliveryReportEvidence.MessageClass)throw new ArchiveReviewException("Avisos de falha exigem a cópia nativa; conversão por MSG bloqueada para preservar o relatório.");
        AttachVolume(volume);
        object? folder = null, staged = null, imported = null;
        string working = Path.Combine(Path.GetDirectoryName(item.StageMsgPath!)!, "w" + Guid.NewGuid().ToString("N")[..16] + ".msg");
        try
        {
            folder = TargetFolder(volume, item.RelativePath, true)!;
            File.Copy(item.StageMsgPath!, working, false); // Immutable backup is never tagged or saved by Outlook.
            using(ArchiveOperationTiming.Measure("import.open-shared-msg"))staged = ((dynamic)session!).OpenSharedItem(working);
            SetProperty(staged, TokenProperty, item.DestinationToken); // Local MSG object only, never an original.
            using(ArchiveOperationTiming.Measure("import.move-copy-to-pst"))imported = ((dynamic)staged).Move(folder);
            var identity = new DestinationIdentity(volume.StoreId!, (string)((dynamic)folder).EntryID, (string)((dynamic)imported).EntryID);
            object? importedParent = null;
            try
            {
                importedParent = ((dynamic)imported).Parent;
                if ((string)((dynamic)importedParent).StoreID != volume.StoreId || (string)((dynamic)importedParent).EntryID != identity.FolderId || identity.EntryId == item.SourceEntryId)
                    throw new ArchiveReviewException("Gravação de identidade bloqueada fora da cópia de destino.");
                SetProperty(imported, TokenProperty, item.DestinationToken);
                using(ArchiveOperationTiming.Measure("import.save-token"))((dynamic)imported).Save(); // Commit the STORED copy.
            }
            finally { Release(importedParent); }
            if (tokenIndexes.TryGetValue(volume.PstPath, out var index))
            {
                if ((string?)Property(imported, TokenProperty) == item.DestinationToken)
                {
                    if (!index.Tokens.TryGetValue(item.DestinationToken, out var matches)) index.Tokens[item.DestinationToken] = matches = [];
                    matches.Add(identity);
                }
                else index.Unmarked++;
            }
            return identity;
        }
        finally
        {
            Release(imported); Release(staged); Release(folder);
            try { File.Delete(working); } catch (IOException) { /* An orphan working copy is not the authoritative backup. */ }
        }
    }
    public ArchiveSnapshot TargetSnapshot(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity, string evidenceDirectory, bool requireToken = true)
    {
        using var timing=ArchiveOperationTiming.Measure("target.snapshot");
        AttachVolume(volume); object? copy = null, parent = null;
        try
        {
            copy = ((dynamic)session!).GetItemFromID(identity.EntryId, volume.StoreId);
            parent = ((dynamic)copy).Parent;
            if ((string)((dynamic)parent).StoreID != volume.StoreId || (string)((dynamic)parent).EntryID != identity.FolderId ||
                (requireToken && (string?)Property(copy, TokenProperty) != item.DestinationToken)) throw new ArchiveReviewException("Identidade da cópia não confere.");
            return Snapshot(copy, evidenceDirectory, RetainComparisonSnapshots, item.RepairToken is not null);
        }
        finally { Release(parent); Release(copy); }
    }
    public List<DestinationIdentity> UnmarkedCandidates(ArchiveVolume volume, ArchiveItem item)
    {
        if (volume.PreserveFolderStructure) return [];
        AttachVolume(volume); object? folder = null, table = null, columns = null;
        var values = new List<DestinationIdentity>();
        try
        {
            folder = TargetFolder(volume, item.RelativePath, false); if (folder is null) return values;
            string folderId = (string)((dynamic)folder).EntryID;
            table = ((dynamic)folder).GetTable("", 0); columns = ((dynamic)table).Columns;
            ((dynamic)columns).RemoveAll(); ((dynamic)columns).Add("EntryID"); ((dynamic)columns).Add(TokenProperty + "/0x0000001f");
            int expected = (int)((dynamic)table).GetRowCount(), scanned = 0;
            var observed = new HashSet<string>(StringComparer.Ordinal);
            while (!(bool)((dynamic)table).EndOfTable)
            {
                Array rows = (Array)((dynamic)table).GetArray(250);
                for (int i = rows.GetLowerBound(0); i <= rows.GetUpperBound(0); i++)
                {
                    string? entryId = rows.GetValue(i, rows.GetLowerBound(1)) as string;
                    object? token = rows.GetValue(i, rows.GetLowerBound(1) + 1);
                    if (string.IsNullOrWhiteSpace(entryId) || !observed.Add(entryId)) throw new ArchiveReviewException("Inventário de candidatos sem identidade única.");
                    if (token is null || token is string text && text.Length == 0)
                        values.Add(new(volume.StoreId!, folderId, entryId));
                    else if (token is not string) throw new ArchiveReviewException("Identificador de candidato não pôde ser lido; revisão necessária.");
                    scanned++;
                }
            }
            if (scanned != expected || (int)((dynamic)table).GetRowCount() != expected) throw new ArchiveReviewException("PST mudou durante o inventário de candidatos.");
            return values;
        }
        finally { Release(columns); Release(table); Release(folder); }
    }
    public void MarkCopy(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity)
    {
        AttachVolume(volume); object? copy = null, parent = null, expected = null;
        try
        {
            expected = TargetFolder(volume, item.RelativePath, false) ?? throw new ArchiveReviewException("Pasta planejada ausente.");
            copy = ((dynamic)session!).GetItemFromID(identity.EntryId, volume.StoreId); parent = ((dynamic)copy).Parent;
            if ((string)((dynamic)parent).StoreID != volume.StoreId || (string)((dynamic)parent).EntryID != (string)((dynamic)expected).EntryID ||
                identity.EntryId == item.SourceEntryId) throw new ArchiveReviewException("Candidato não pertence ao destino planejado.");
            string? old = Property(copy, TokenProperty) as string;
            if (!string.IsNullOrEmpty(old) && old != item.DestinationToken) throw new ArchiveReviewException("Candidato já pertence a outro item.");
            SetProperty(copy, TokenProperty, item.DestinationToken); ((dynamic)copy).Save();
            if ((string?)Property(copy, TokenProperty) != item.DestinationToken) throw new ArchiveReviewException("Identificador não persistiu após Save.");
            tokenIndexes.Remove(volume.PstPath);
        }
        finally { Release(expected); Release(parent); Release(copy); }
    }
    public DiskRoundtrip ReopenVolume(ArchiveVolume volume)
    {
        AttachVolume(volume); object? store = null, root = null; bool exclusive = false, detached = false;
        try
        {
            store = FindStore(volume.PstPath, volume.StoreId)!;
            object? inspectors = null, explorers = null;
            try
            {
                inspectors = ((dynamic)app!).Inspectors;
                // An unrelated mailbox editor does not hold this PST. Never
                // close user windows; block only a target-store or unknown editor.
                for (int i = 1; i <= (int)((dynamic)inspectors).Count; i++)
                {
                    object? inspector = null, current = null, parent = null;
                    try
                    {
                        inspector = ((dynamic)inspectors)[i]; current = ((dynamic)inspector).CurrentItem;
                        parent = ((dynamic)current).Parent;
                        string currentStore = (string)((dynamic)parent).StoreID;
                        if (string.IsNullOrWhiteSpace(currentStore) || OutlookIdentity.AreSame(currentStore, volume.StoreId!,
                            (left, right) => (bool)((dynamic)session!).CompareEntryIDs(left, right)))
                            throw new ArchiveReviewException("Há uma mensagem deste PST aberta; feche essa mensagem para validar o arquivo.");
                    }
                    catch (COMException) { throw new ArchiveReviewException("Há um editor cuja origem não pôde ser identificada; feche essa janela para validar o PST."); }
                    finally { Release(parent); Release(current); Release(inspector); }
                }
                explorers = ((dynamic)app!).Explorers;
                for (int i = 1; i <= (int)((dynamic)explorers).Count; i++)
                {
                    object? explorer = null, folder = null;
                    try
                    {
                        explorer = ((dynamic)explorers)[i]; folder = ((dynamic)explorer).CurrentFolder;
                        if ((string)((dynamic)folder).StoreID == volume.StoreId) throw new ArchiveReviewException("PST em exibição numa janela; selecione outra caixa para validar.");
                    }
                    finally { Release(folder); Release(explorer); }
                }
            }
            finally { Release(explorers); Release(inspectors); }
            root = ((dynamic)store).GetRootFolder(); detached = true;
            attached.Remove(volume.PstPath); tokenIndexes.Remove(volume.PstPath); folderIds.Clear();
            ((dynamic)session!).RemoveStore(root);
            Release(root); root = null; Release(store); store = null;
            for (int i = 0; i < 5; i++)
            {
                try { using var input = new FileStream(volume.PstPath, FileMode.Open, FileAccess.Read, FileShare.None); exclusive = input.Length > 0; break; }
                catch (IOException) { if (i < 4) Thread.Sleep(200); }
            }
            AttachVolume(volume); detached = false;
            return new(true, exclusive);
        }
        finally
        {
            Release(root); Release(store);
            if (detached) AttachVolume(volume);
        }
    }
    public SourceLocation? LocateSource(ArchiveItem item)
    {
        object? source = null, parent = null, trash = null, originalFolder=null;
        try
        {
            string id = item.RemovalEntryId ?? item.SourceEntryId;
            source = TryItem(id, item.SourceStoreId);
            trash = DeletedFolder(item.SourceStoreId);
            string trashId = (string)((dynamic)trash).EntryID;
            if(source is null && item.SourceRecordKey is not null && item.RemovalEntryId is null)
            {
                try{originalFolder=((dynamic)session!).GetFolderFromID(item.SourceFolderId,item.SourceStoreId);source=FindIndexed(originalFolder,item.SourceRecordKey);}
                catch(COMException ex) when(ex.HResult==unchecked((int)0x8004010F)){}
            }
            if (source is null && item.SourceRecordKey is not null)
                source = FindIndexed(trash, item.RemovalRecordKey ?? item.SourceRecordKey);
            bool recoveredByMarker = false;
            if (source is null && item.State is ArchiveItemState.DeleteStarted or ArchiveItemState.InDeletedItems or ArchiveItemState.PurgeStarted)
            {
                source = FindIndexed(trash, item.Token,true);
                recoveredByMarker = source is not null;
            }
            if (source is null) return null;
            parent = ((dynamic)source).Parent;
            string folderId = (string)((dynamic)parent).EntryID;
            if ((string)((dynamic)parent).StoreID != item.SourceStoreId || (folderId != item.SourceFolderId && folderId != trashId))
                throw new ArchiveReviewException("Origem foi movida externamente para outra pasta.");
            string? key = BinaryProperty(source, "0x0FF90102");
            string? expected = item.RemovalEntryId is null ? item.SourceRecordKey : item.RemovalRecordKey;
            if (!recoveredByMarker && expected is not null && key != expected) throw new ArchiveReviewException("RecordKey da origem mudou.");
            return new((string)((dynamic)source).EntryID, folderId, folderId == trashId, key);
        }
        finally { Release(originalFolder);Release(trash); Release(parent); Release(source); }
    }
    public SourceLocation MoveToDeletedItems(ArchiveItem item, string? expectedChangeKey)
    {
        object? original = null, trash = null, moved = null;
        try
        {
            var current=LocateSource(item)??throw new ArchiveReviewException("Original não localizado antes do movimento.");
            original = OpenSource(item,current.EntryId,current.FolderId,current.RecordKey);
            CheckSourceRemovalMode(original);
            CheckChangeKey(original, expectedChangeKey);
            // PST Move changes EntryID AND RecordKey. Persist a unique recovery
            // marker only after the engine has committed its validated removal
            // intent. Never identify an original by subject/content alone.
            string? oldMarker = Property(original, RemovalTokenProperty) as string;
            if (oldMarker is not null && oldMarker != item.Token) throw new ArchiveReviewException("Original já vinculado a outra remoção.");
            SetProperty(original, RemovalTokenProperty, item.Token);
            ((dynamic)original).Save();
            var marked = Snapshot(original, Path.Combine(Path.GetDirectoryName(item.StageMsgPath!)!, "mark-" + Guid.NewGuid().ToString("N")[..12]), false, item.RepairToken is not null);
            ArchiveSnapshot? replay = null;
            if (EmbeddedMessageEvidence.Indices(marked.Fields).Count > 0)
                replay = BackupSnapshot(item, Path.Combine(Path.GetDirectoryName(item.StageMsgPath!)!, "mark-backup-" + Guid.NewGuid().ToString("N")[..12]));
            if (!ArchivedSourceEvidence.Matches(item, marked, replay)) throw new ArchiveReviewException("Informações do original mudaram ao preparar a remoção; movimentação bloqueada.");
            CheckChangeKey(original, marked.ChangeKey);
            trash = DeletedFolder(item.SourceStoreId);
            string trashId = (string)((dynamic)trash).EntryID;
            if (trashId == current.FolderId) return current;
            moved = ((dynamic)original).Move(trash);
            var location=new SourceLocation((string)((dynamic)moved).EntryID,trashId,true,BinaryProperty(moved,"0x0FF90102"));
            UpdateSourceIndex(item,location,current.EntryId);return location;
        }
        finally { Release(moved); Release(trash); Release(original); }
    }
    private static object? FindRemovalToken(object folder, string token)
    {
        object? items = null, found = null;
        try
        {
            items = ((dynamic)folder).Items;
            for (int i = 1; i <= (int)((dynamic)items).Count; i++)
            {
                object? candidate = ((dynamic)items)[i];
                if (Property(candidate, RemovalTokenProperty) as string == token)
                {
                    if (found is not null) { Release(candidate); throw new ArchiveReviewException("Identificador de remoção ambíguo em Itens Excluídos."); }
                    found = candidate;
                }
                else Release(candidate);
            }
            object? result = found; found = null; return result;
        }
        finally { Release(found); Release(items); }
    }
    public void DeleteFromDeletedItems(ArchiveItem item, SourceLocation location, string? expectedChangeKey)
    {
        object? trash = null, original = null;
        try
        {
            trash = DeletedFolder(item.SourceStoreId);
            if (!location.InDeletedItems || (string)((dynamic)trash).EntryID != location.FolderId) throw new ArchiveReviewException("Remoção só é permitida no item rastreado em Itens Excluídos.");
            original = OpenSource(item, location.EntryId, location.FolderId, location.RecordKey);
            CheckSourceRemovalMode(original);
            CheckChangeKey(original, expectedChangeKey);
            ((dynamic)original).Delete(); // One tracked, validated item; NEVER EmptyFolder.
            if(sourceLookupIndexes.TryGetValue(LookupKey(item.SourceStoreId,location.FolderId),out var cached))cached.Index.Remove(location.EntryId);
        }
        finally { Release(original); Release(trash); }
    }
    private object OpenSource(ArchiveItem item, string entryId, string folderId, string? recordKey)
    {
        object? source = TryItem(entryId, item.SourceStoreId); object? parent = null;
        if (source is null) throw new ArchiveReviewException("Original não localizado pelo identificador registrado; ele pode ter sido movido ou excluído. Remoção bloqueada.");
        try
        {
            parent = ((dynamic)source).Parent;
            if ((string)((dynamic)parent).StoreID != item.SourceStoreId || (string)((dynamic)parent).EntryID != folderId ||
                !DeliveryReportEvidence.SupportedSource(source) ||
                (recordKey is not null && BinaryProperty(source, "0x0FF90102") != recordKey)) throw new ArchiveReviewException("Origem mudou, é rascunho ou está fora do escopo.");
            return source;
        }
        catch { Release(source); throw; }
        finally { Release(parent); }
    }
    private object DeletedFolder(string storeId)
    {
        object? store = null;
        try
        {
            store = ((dynamic)session!).GetStoreFromID(storeId);
            return ((dynamic)store).GetDefaultFolder(3)
                ?? throw new ArchiveReviewException("Itens Excluídos não disponíveis nesta origem.");
        }
        finally { Release(store); }
    }
    private object? TargetFolder(ArchiveVolume volume, string[] path, bool create)
    {
        string cacheKey = volume.JobId + ":" + volume.Number + ":" + System.Text.Json.JsonSerializer.Serialize(path);
        if (folderIds.TryGetValue(cacheKey, out string? cachedId)) return ((dynamic)session!).GetFolderFromID(cachedId, volume.StoreId);
        object? store = null, current = null;
        try
        {
            store = FindStore(volume.PstPath, volume.StoreId) ?? throw new ArchiveReviewException("PST indisponível."); current = ((dynamic)store).GetRootFolder();
            if (volume.PreserveFolderStructure)
            {
                Release(current); current = StructuredTargetFolder(store, volume, path, create);
                if (current is null) return null;
                folderIds[cacheKey] = (string)((dynamic)current).EntryID;
                Console.WriteLine($"Pasta de destino: {string.Join(" > ", path)} → {(string)((dynamic)current).FolderPath}");
                object mapped = current; current = null; return mapped;
            }
            foreach (string segment in new[] { volume.IsExisting ? "Arquivo-" + volume.JobId : "Arquivo" }.Concat(path))
            {
                object? folders = null, found = null;
                try
                {
                    folders = ((dynamic)current).Folders;
                    for (int i = 1; i <= (int)((dynamic)folders).Count; i++)
                    {
                        object? candidate = ((dynamic)folders)[i];
                        if ((string)((dynamic)candidate).Name == segment)
                        {
                            if (found is not null) { Release(candidate); throw new ArchiveReviewException("Pasta de destino ambígua."); }
                            found = candidate;
                        }
                        else Release(candidate);
                    }
                    if (found is null && create) found = ((dynamic)folders).Add(segment, 6);
                    if (found is null) return null;
                    Release(current); current = found; found = null;
                }
                finally { Release(found); Release(folders); }
            }
            folderIds[cacheKey] = (string)((dynamic)current).EntryID;
            object result = current; current = null; return result;
        }
        finally { Release(current); Release(store); }
    }
    private object? FindStore(string path, string? recordedStoreId = null)
    {
        string? knownId = attached.GetValueOrDefault(path) ?? recordedStoreId;
        if (!string.IsNullOrWhiteSpace(knownId))
        {
            var known = RecordedStoreLookup.Open(path, knownId,
                id => ((dynamic)session!).GetStoreFromID(id),
                store => (string)((dynamic)store).FilePath, Release);
            if (known is not null) return known;
            attached.Remove(path);
        }
        string? discoveredId = ResolveStoreIdByPath(path) ?? ResolveStoreIdFromOutlookMetadata(path);
        return discoveredId is null ? null : RecordedStoreLookup.Open(path, discoveredId,
            id => ((dynamic)session!).GetStoreFromID(id), store => (string)((dynamic)store).FilePath, Release);
    }
    private object? TryItem(string id, string store)
    {
        try { return ((dynamic)session!).GetItemFromID(id, store); }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8004010F)) { return null; }
    }
    private object? FindByRecordKey(object folder, string recordKey)
    {
        object? table = null, columns = null;
        try
        {
            table = ((dynamic)folder).GetTable("", 0); columns = ((dynamic)table).Columns;
            ((dynamic)columns).RemoveAll(); ((dynamic)columns).Add("EntryID");
            ((dynamic)columns).Add("http://schemas.microsoft.com/mapi/proptag/0x0FF90102");
            string? foundId = null;
            while (!(bool)((dynamic)table).EndOfTable)
            {
                Array rows = (Array)((dynamic)table).GetArray(100);
                for (int i = rows.GetLowerBound(0); i <= rows.GetUpperBound(0); i++)
                {
                    object? key = rows.GetValue(i, rows.GetLowerBound(1) + 1);
                    if (key is not byte[] bytes || Convert.ToHexString(bytes) != recordKey) continue;
                    if (foundId is not null) throw new ArchiveReviewException("RecordKey ambígua; remoção bloqueada.");
                    foundId = rows.GetValue(i, rows.GetLowerBound(1)) as string;
                }
            }
            return foundId is null ? null : TryItem(foundId, (string)((dynamic)folder).StoreID);
        }
        finally { Release(columns); Release(table); }
    }
    private static string Owner(ArchiveVolume volume) => volume.JobId + ":" + volume.Number;
    private static object? Property(object item, string name)
    {
        object? accessor = null;
        try { accessor = ((dynamic)item).PropertyAccessor; return ((dynamic)accessor).GetProperty(name); }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8004010F)) { return null; }
        finally { Release(accessor); }
    }
    private static void SetProperty(object item, string name, string value)
    {
        object? accessor = null;
        try { accessor = ((dynamic)item).PropertyAccessor; ((dynamic)accessor).SetProperty(name, value); }
        finally { Release(accessor); }
    }
    private static string? BinaryProperty(object item, string tag) => Property(item, "http://schemas.microsoft.com/mapi/proptag/" + tag) is byte[] bytes ? Convert.ToHexString(bytes) : null;
    private static void CheckChangeKey(object item, string? expected)
    { if (BinaryProperty(item, "0x65E20102") != expected) throw new ArchiveReviewException("Mensagem mudou imediatamente antes da remoção."); }
    private void CheckSourceRemovalMode(object item)
    {
        object? parent = null, store = null;
        try
        {
            parent = ((dynamic)item).Parent; store = ((dynamic)parent).Store;
            string path = (string)((dynamic)store).FilePath;
            if (!string.IsNullOrEmpty(path) && string.Equals(Path.GetExtension(path), ".pst", StringComparison.OrdinalIgnoreCase))
            { ArchiveFiles.ValidateLocalDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!); return; }
            if ((bool)((dynamic)session!).Offline || (bool)((dynamic)store).IsCachedExchange)
                throw new ArchiveReviewException("Remoção de Exchange offline/em cache bloqueada: esta versão exige origem on-line sem cache para não validar apenas uma cópia local desatualizada.");
        }
        finally { Release(store); Release(parent); }
    }
    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    public void Dispose() { Release(session); Release(app); session = null; app = null; foreach (var lease in pstLeases.Values) lease.Dispose(); pstLeases.Clear(); }
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? instance);
}
