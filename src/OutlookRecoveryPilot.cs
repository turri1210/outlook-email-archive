using System.Text.Json;

internal sealed partial class OutlookArchiveBackend
{
    internal void InspectFastRecovery(string rootDirectory, string jobId, string key)
    {
        using var lease = new ArchiveEngine(rootDirectory, this);
        using var ledger = new ArchiveLedger(Path.Combine(rootDirectory, "archive.sqlite"), true);
        var job = ledger.GetJob(jobId); var row = ledger.GetItems(jobId).Single(x => x.Key == key);
        if (row.FastIdentityPath is null) throw new ArchiveReviewException("Identidade rápida ausente.");
        var initial = ArchiveFiles.ReadFields(ArchiveFiles.Child(job.RootDirectory, row.FastIdentityPath));
        if (ArchiveFiles.HashFields(initial) != row.FastIdentityHash) throw new ArchiveReviewException("Identidade rápida alterada.");
        var current = CaptureSourceIdentity(row);
        var changes = initial.Keys.Union(current.Fields.Keys).Where(field => initial.GetValueOrDefault(field) != current.Fields.GetValueOrDefault(field))
            .Select(field => new { Field = field, Initial = field is "Sent" or "Received" ? initial.GetValueOrDefault(field) : "omitido", Current = field is "Sent" or "Received" ? current.Fields.GetValueOrDefault(field) : "omitido" }).ToArray();
        RefreshSourceIdentities(new[] { row });
        var native = CaptureSourceIdentity(row);
        var nativeChanges = initial.Keys.Union(native.Fields.Keys).Where(field => initial.GetValueOrDefault(field) != native.Fields.GetValueOrDefault(field)).ToArray();
        Console.WriteLine(JsonSerializer.Serialize(new { Key = key, Changes = changes, NativeChanges = nativeChanges, NativeExactlyMatches = native.Hash == row.FastIdentityHash }, ArchiveFiles.JsonOptions));
    }
    internal void InspectRecoveryEvidence(string rootDirectory, string jobId, string key)
    {
        using var lease = new ArchiveEngine(rootDirectory, this);
        using var ledger = new ArchiveLedger(Path.Combine(rootDirectory, "archive.sqlite"), true);
        var job = ledger.GetJob(jobId); var row = ledger.GetItems(jobId).Single(x => x.Key == key);
        if (row.State != ArchiveItemState.ReviewRequired || row.TargetEntryId is null || row.StageMsgPath is null)
            throw new ArchiveReviewException("Inspeção exige item em revisão com backup e cópia registrados.");
        var volume = ledger.GetVolumes(jobId).Single(x => x.Number == row.VolumeNumber);
        if (ArchiveFiles.Sha(row.StageMsgPath) != row.StageMsgSha256) throw new ArchiveReviewException("Backup alterado.");
        string directory = Path.Combine(job.RootDirectory, "recovery-analysis", key[..16] + "-" + Guid.NewGuid().ToString("N")[..8]);
        var source = Export(row, Path.Combine(directory, "source"));
        AttachVolume(volume); object? target = null, parent = null;
        try
        {
            target = ((dynamic)session!).GetItemFromID(row.TargetEntryId, volume.StoreId); parent = ((dynamic)target).Parent;
            if ((string)((dynamic)parent).StoreID != volume.StoreId || (string)((dynamic)parent).EntryID != row.TargetFolderId ||
                (string?)Property(target, TokenProperty) != row.Token) throw new ArchiveReviewException("Destino não corresponde ao registro.");
            var snapshot = Snapshot(target, Path.Combine(directory, "target"), true);
            ArchiveFiles.WriteNew(Path.Combine(directory, "inspection.json"), new { Key = key, Evidence = directory,
                Equal = RelevantMessageValidation.Equivalent(source, snapshot), Differences = RelevantMessageValidation.DifferenceSummary(source, snapshot),
                OriginalDeleted = false, LedgerChanged = false });
            Console.WriteLine(directory);
        }
        finally { Release(parent); Release(target); }
    }
    internal void CleanupRecoveryPilot(string rootDirectory, string jobId, string key)
    {
        using var lease = new ArchiveEngine(rootDirectory, this);
        using var ledger = new ArchiveLedger(Path.Combine(rootDirectory, "archive.sqlite"), true);
        var job = ledger.GetJob(jobId);
        var row = ledger.GetItems(jobId).Single(x => x.Key == key);
        var volume = ledger.GetVolumes(jobId).Single(x => x.Number == row.VolumeNumber);
        string evidence = Path.Combine(job.RootDirectory, "recovery-pilot", key[..16]);
        var identity = JsonSerializer.Deserialize<DestinationIdentity>(File.ReadAllText(Path.Combine(evidence, "destination.json")))!;
        string name = "Recuperação controlada-" + key[..12];
        if (identity.StoreId != volume.StoreId || identity.EntryId == row.SourceEntryId || identity.EntryId == row.TargetEntryId)
            throw new ArchiveReviewException("Limpeza fora da identidade de teste bloqueada.");
        object? folder = null, parent = null, store = null, owner = null, items = null, child = null, trash = null;
        try
        {
            AttachVolume(volume); store = FindStore(volume.PstPath, volume.StoreId)!; owner = OwnedScope(store, volume, false);
            folder = ((dynamic)session!).GetFolderFromID(identity.FolderId, identity.StoreId);
            parent = ((dynamic)folder).Parent;
            items = ((dynamic)folder).Items;
            if ((string)((dynamic)folder).Name != name || (string)((dynamic)parent).EntryID != (string)((dynamic)owner).EntryID ||
                (int)((dynamic)items).Count != 1 || (int)((dynamic)folder).Folders.Count != 0)
                throw new ArchiveReviewException("Pasta de teste mudou; limpeza bloqueada.");
            child = ((dynamic)items)[1];
            if ((string)((dynamic)child).EntryID != identity.EntryId) throw new ArchiveReviewException("Conteúdo de teste mudou; limpeza bloqueada.");
            Release(child); child = null; Release(items); items = null; Release(parent); parent = null;
            ((dynamic)folder).Delete(); Release(folder); folder = null;
            // Reopen the same folder identity. Only purge if Outlook confirms
            // that this exact isolated test folder is now in this PST's trash.
            folder = ((dynamic)session!).GetFolderFromID(identity.FolderId, identity.StoreId);
            parent = ((dynamic)folder).Parent; trash = ((dynamic)store).GetDefaultFolder(3);
            items = ((dynamic)folder).Items; child = ((dynamic)items)[1];
            if ((string)((dynamic)parent).EntryID != (string)((dynamic)trash).EntryID || (string)((dynamic)folder).Name != name ||
                (int)((dynamic)items).Count != 1 || (string)((dynamic)child).EntryID != identity.EntryId || (int)((dynamic)folder).Folders.Count != 0)
                throw new ArchiveReviewException("Destino da pasta de teste não confirmado; limpeza final bloqueada.");
            Release(child); child = null; Release(items); items = null;
            ((dynamic)folder).Delete();
            ArchiveFiles.WriteNew(Path.Combine(evidence, "cleanup.json"), new { TestFolderRemoved = true, OriginalRemoved = false, PreviousCopyRemoved = false, Utc = DateTimeOffset.UtcNow });
            Console.WriteLine("Cópia extra e pasta de teste removidas. Original, cópia registrada e arquivo PST preservados.");
        }
        finally { Release(child); Release(items); Release(trash); Release(parent); Release(folder); Release(owner); Release(store); }
    }
    // One controlled copy, outside the migration ledger. No source deletion,
    // baseline replacement or automatic repetition after an interrupted Move.
    internal int RecoveryPilot(string rootDirectory, string jobId, string key)
    {
        using var lease = new ArchiveEngine(rootDirectory, this);
        using var ledger = new ArchiveLedger(Path.Combine(rootDirectory, "archive.sqlite"), true);
        var job = ledger.GetJob(jobId);
        var row = ledger.GetItems(jobId).Single(x => x.Key == key);
        if (row.State != ArchiveItemState.ReviewRequired || row.StageMsgPath is null || row.TargetEntryId is null)
            throw new ArchiveReviewException("O piloto exige item em revisão com backup e destino registrados.");
        var volume = ledger.GetVolumes(jobId).Single(x => x.Number == row.VolumeNumber);
        if (ArchiveFiles.Sha(row.StageMsgPath) != row.StageMsgSha256)
            throw new ArchiveReviewException("Backup original alterado; piloto bloqueado.");
        string evidence = Path.Combine(job.RootDirectory, "recovery-pilot", row.Key[..16]);
        Directory.CreateDirectory(evidence);
        string folderName = "Recuperação controlada-" + row.Key[..12];
        ArchiveFiles.WriteNew(Path.Combine(evidence, "intent.json"), new { jobId, key, volume.PstPath, Folder = folderName, Utc = DateTimeOffset.UtcNow });
        Console.WriteLine("Piloto: conferindo original e backup antes de criar uma cópia…");
        var backup = BackupSnapshot(row, Path.Combine(evidence, "backup"));
        var before = Export(row, Path.Combine(evidence, "source-before"));
        if (backup.Hash != row.SnapshotHash || !RelevantMessageValidation.Equivalent(backup, before))
            throw new ArchiveReviewException("Original e backup não correspondem; nenhuma cópia foi criada.");
        var previous = TargetSnapshot(volume, row, new(row.TargetStoreId!, row.TargetFolderId!, row.TargetEntryId), Path.Combine(evidence, "previous-copy"));
        AttachVolume(volume);
        object? store = null, owner = null, folders = null, folder = null, original = null, staged = null, copied = null;
        try
        {
            store = FindStore(volume.PstPath, volume.StoreId) ?? throw new ArchiveReviewException("Destino registrado ausente.");
            owner = OwnedScope(store, volume, false); folders = ((dynamic)owner).Folders;
            folder = ((dynamic)folders).Add(folderName, 6);
            ArchiveFiles.WriteNew(Path.Combine(evidence, "folder.json"), new { StoreId = volume.StoreId, FolderId = (string)((dynamic)folder).EntryID });
            original = OpenSource(row, row.SourceEntryId, row.SourceFolderId, row.SourceRecordKey);
            CheckChangeKey(original, before.ChangeKey);
            Console.WriteLine("Piloto: criando cópia pelo Outlook e transferindo apenas essa cópia para o PST…");
            staged = ((dynamic)original).Copy();
            string stagedId = (string)((dynamic)staged).EntryID;
            if (stagedId == row.SourceEntryId) throw new ArchiveReviewException("Outlook não criou uma identidade distinta; movimentação bloqueada.");
            ArchiveFiles.WriteNew(Path.Combine(evidence, "source-copy-checkpoint.json"), new { StoreId = row.SourceStoreId, EntryId = stagedId, OriginalEntryId = row.SourceEntryId });
            copied = ((dynamic)staged).Move(folder);
            var identity = new DestinationIdentity(volume.StoreId!, (string)((dynamic)folder).EntryID, (string)((dynamic)copied).EntryID);
            ArchiveFiles.WriteNew(Path.Combine(evidence, "destination.json"), identity);
            var target = Snapshot(copied, Path.Combine(evidence, "new-copy"), true);
            var after = SourceSnapshot(row, null, Path.Combine(evidence, "source-after"));
            bool sourceIntact = RelevantMessageValidation.Equivalent(before, after);
            bool backupEqual = RelevantMessageValidation.Equivalent(backup, target);
            bool currentEqual = RelevantMessageValidation.Equivalent(before, target);
            bool approved = sourceIntact && backupEqual && currentEqual;
            var result = new { Approved = approved, SourceIntact = sourceIntact, BackupEqual = backupEqual, CurrentSourceEqual = currentEqual,
                PreviousCopyEqual = RelevantMessageValidation.Equivalent(backup, previous),
                PreviousDifferences = RelevantMessageValidation.DifferenceSummary(backup, previous),
                NewDifferences = approved ? "Nenhuma" : RelevantMessageValidation.DifferenceSummary(backup, target),
                Destination = identity, Evidence = evidence, OriginalDeleted = false, LedgerChanged = false };
            ArchiveFiles.WriteNew(Path.Combine(evidence, "result.json"), result);
            Console.WriteLine(JsonSerializer.Serialize(result, ArchiveFiles.JsonOptions));
            return approved ? 0 : 3;
        }
        finally { Release(copied); Release(staged); Release(original); Release(folder); Release(folders); Release(owner); Release(store); }
    }
}
