using System.Text.Json;

internal sealed partial class ArchiveEngine
{
    private bool NativeWorkersIdle(string jobId)
    {
        if (backend is not IArchiveBatchBackend native) return true;
        try { foreach (var batch in ledger.GetBatches(jobId)) native.AssertBatchIdle(batch); return true; }
        catch (IOException)
        { Console.Error.WriteLine("O Outlook ainda está concluindo um lote interrompido. Aguarde e retome esta execução; o PST não será fechado durante a cópia."); return false; }
    }
    private sealed record BatchMember(string Key, string EntryId, string SnapshotHash, string MsgSha256, string? RelevantHash, long Reservation);
    private sealed record BatchManifest(string Id, string JobId, int VolumeNumber, string SourceStoreId, string SourceFolderId,
        string[] RelativePath, List<BatchMember> Members, bool FastCopy = false);

    private bool CopyNativeBatches(ArchiveJob job, ArchiveConfig config, IArchiveBatchBackend native)
    {
        var bulk = backend as IArchiveBulkFastBatchBackend;
        bulk?.ClearFastIdentityCache();
        try { return CopyNativeBatchesCore(job, config, native); }
        finally { bulk?.ClearFastIdentityCache(); }
    }
    private bool CopyNativeBatchesCore(ArchiveJob job, ArchiveConfig config, IArchiveBatchBackend native)
    {
        // Always reconcile existing batches before forming new ones. No member
        // is sent through the old per-message importer after durable linking.
        foreach (var batch in ledger.GetBatches(job.Id).Where(x => (workflowSelection is null || x.Items.Any(workflowSelection.Contains)) && x.State is not (ArchiveBatchState.Complete or ArchiveBatchState.ReviewRequired)))
        {
            if (!RunNativeBatch(job, config, native, batch) && batch.State != ArchiveBatchState.ReviewRequired) return false;
        }
        var pending = new List<ArchiveItem>(); ArchiveVolume? pendingVolume = null;
        var candidates = WorkItems(job.Id).Where(x => x.NativeBatchId is null && x.State is ArchiveItemState.Planned or ArchiveItemState.Exported).ToList();
        for (int index = 0; index < candidates.Count; index++)
        {
            var row = candidates[index];
            if (index % 100 == 0 && backend is IArchiveBulkFastBatchBackend bulk)
            {
                Console.WriteLine($"Lendo identificações em lote: {index:N0}/{candidates.Count:N0} itens do escopo inicial…");
                var ordinary=candidates.Skip(index).Take(100).Where(x=>x.SourceMessageClass is null or "IPM.Note").ToList();
                if(ordinary.Count>0)bulk.RefreshSourceIdentities(ordinary);
            }
            try
            {
                if (!native.CanBatch(row)) continue;
                bool fast = backend is IArchiveFastBatchBackend && row.State == ArchiveItemState.Planned && row.SourceMessageClass != DeliveryReportEvidence.MessageClass;
                if (fast)
                {
                    var identity = ((IArchiveFastBatchBackend)backend).CaptureSourceIdentity(row);
                    row.FastIdentityPath = Path.Combine(Evidence(job, row, "fast-identity"), "fields.json");
                    Directory.CreateDirectory(Path.GetDirectoryName(row.FastIdentityPath)!);
                    ArchiveFiles.WriteNew(row.FastIdentityPath, identity.Fields);
                    row.FastIdentityHash = identity.Hash; row.SourceChangeKey = identity.ChangeKey;
                    ledger.UpdateItem(row);
                }
                else if (row.State == ArchiveItemState.Planned)
                {
                    CheckSpace(row.EstimatedBytes, config);
                    var original = backend.Export(row, Evidence(job, row, "original"));
                    row.StageMsgPath = ArchiveFiles.Child(job.RootDirectory, original.MessagePath);
                    row.StageMsgSha256 = original.MessageSha256; row.SnapshotHash = original.Hash; row.SourceChangeKey = original.ChangeKey;
                    SaveRelevantOriginal(row, original); row.State = ArchiveItemState.Exported; row.Error = null;
                    ledger.UpdateItem(row); Fault?.Invoke("after-export", row);
                }
                if (!fast)
                {
                CheckBackupFile(job, row);
                if (row.SourceRelevantHash is null)
                {
                    var backup = backend.BackupSnapshot(row, Evidence(job, row, "native-backup-profile"));
                    if (backup.Hash != row.SnapshotHash || backup.MessageSha256 != row.StageMsgSha256)
                        throw new ArchiveReviewException("Backup original não corresponde à evidência registrada.");
                    SaveRelevantOriginal(row, backup); ledger.UpdateItem(row);
                }
                var snapshot = OriginalSnapshot(job, row);
                if ((row.SourceMessageClass != DeliveryReportEvidence.MessageClass && !RelevantMessageValidation.Supported(snapshot)) || new FileInfo(row.StageMsgPath!).Length > 64L * 1024 * 1024) continue;
                }
                var volume = Allocate(job, row, config); row.VolumeNumber = volume.Number; ledger.UpdateItem(row);
                bool different = pending.Count > 0 && (volume.Number != pendingVolume!.Number || row.SourceFolderId != pending[0].SourceFolderId ||
                    !row.RelativePath.SequenceEqual(pending[0].RelativePath) || (row.FastIdentityHash is null) != (pending[0].FastIdentityHash is null) || pending.Count >= (fast ? 100 : 50) ||
                    pending.Sum(BatchEstimate) + BatchEstimate(row) > (fast ? 128L : 64L) * 1024 * 1024);
                if (different)
                {
                    if (!StartNativeBatch(job, config, native, pendingVolume!, pending)) return false;
                    pending = [];
                }
                pending.Add(row); pendingVolume = volume;
                Console.WriteLine($"Preparando lote: {pending.Count}/{(fast ? 100 : 50)} mensagens do escopo inicial.");
            }
            catch (ArchiveInjectedCrashException) { throw; }
            catch (ArchiveReviewException ex) { Review(row, ex.Message); }
            catch (Exception ex)
            {
                row.Error = "native-preparation-interrupted:" + ex.GetType().Name + ":" + ex.Message; ledger.UpdateItem(row);
                Console.Error.WriteLine("Preparação interrompida; backups e progresso preservados."); Console.Error.WriteLine(ex.ToString()); return false;
            }
        }
        return pending.Count == 0 || StartNativeBatch(job, config, native, pendingVolume!, pending);
    }
    private bool StartNativeBatch(ArchiveJob job, ArchiveConfig config, IArchiveBatchBackend native, ArchiveVolume volume, List<ArchiveItem> rows)
    {
        var batch = new ArchiveBatch { JobId = job.Id, VolumeNumber = volume.Number, SourceStoreId = rows[0].SourceStoreId,
            SourceFolderId = rows[0].SourceFolderId, RelativePath = rows[0].RelativePath.ToArray(), FastCopy = rows[0].FastIdentityHash is not null, Items = rows.Select(x => x.Key).ToList() };
        batch.DirectoryPath = Path.Combine(job.RootDirectory, "batches", batch.Id);
        Directory.CreateDirectory(batch.DirectoryPath);
        var manifest = new BatchManifest(batch.Id, batch.JobId, batch.VolumeNumber, batch.SourceStoreId, batch.SourceFolderId, batch.RelativePath,
            rows.Select(x => new BatchMember(x.Key, x.SourceEntryId, x.SnapshotHash!, x.StageMsgSha256!, x.FastIdentityHash ?? x.SourceRelevantHash, x.ReservedBytes)).ToList(), batch.FastCopy);
        string manifestPath = Path.Combine(batch.DirectoryPath, "manifest.json"); ArchiveFiles.WriteNew(manifestPath, manifest);
        batch.ManifestSha256 = ArchiveFiles.Sha(manifestPath); ledger.AddBatch(batch, rows);
        return RunNativeBatch(job, config, native, batch) || batch.State == ArchiveBatchState.ReviewRequired;
    }
    private bool RunNativeBatch(ArchiveJob job, ArchiveConfig config, IArchiveBatchBackend native, ArchiveBatch batch)
    {
        var batchClock = System.Diagnostics.Stopwatch.StartNew();
        var rows = batch.Items.Select(x => ledger.GetItem(x) ?? throw new ArchiveReviewException("Membro do lote ausente.")).ToList();
        try
        {
            ValidateBatchManifest(job, batch, rows);
            native.AssertBatchIdle(batch);
            if (rows.All(x=>x.RepairToken is not null || x.TargetEntryId is not null && x.TargetFolderId!=batch.FolderId && x.TargetSnapshotHash is not null && x.TargetFieldsPath is not null))
            {
                foreach(var completed in rows.Where(x=>x.RepairToken is null))_ = TargetBaseline(job,completed);
                batch.State=ArchiveBatchState.Complete;batch.Error=null;ledger.UpdateBatch(batch);
                Console.WriteLine("Lote já finalizado: preservando os destinos registrados e os reparos posteriores.");return true;
            }
            var volume = ledger.GetVolumes(job.Id).Single(x => x.Number == batch.VolumeNumber); EnsureVolume(volume);
            bool copiedThisRun = false;
            if (batch.State == ArchiveBatchState.Prepared)
            {
                var folder = native.PrepareBatch(volume, batch);
                batch.StoreId = folder.StoreId; batch.FolderId = folder.FolderId; ledger.UpdateBatch(batch);
                if (batch.FastCopy && backend is IArchiveBulkFastBatchBackend bulk) bulk.RefreshSourceIdentities(rows);
                foreach (var row in rows) ValidateNativeSource(job, row);
                batch.State = ArchiveBatchState.CopyIntent; ledger.UpdateBatch(batch);
                foreach (var row in rows) { row.State = ArchiveItemState.ImportStarted; row.Error = null; ledger.UpdateItem(row); }
                Fault?.Invoke("after-batch-copy-intent", rows[0]);
                Console.WriteLine($"Copiando lote de {rows.Count} mensagens para a parte {volume.Number}…");
                native.CopyBatch(volume, batch, rows);
                copiedThisRun = true;
                Fault?.Invoke("after-batch-copy-before-checkpoint", rows[0]);
            }
            if (batch.State == ArchiveBatchState.CopyIntent)
            {
                native.AssertBatchIdle(batch);
                if (!copiedThisRun && batch.FastCopy && backend is IArchiveBulkFastBatchBackend refreshed) refreshed.RefreshSourceIdentities(rows);
                var candidates = native.BatchCandidates(volume, batch);
                if (candidates.Count > rows.Count || candidates.Select(x => x.StoreId + ":" + x.EntryId).Distinct().Count() != candidates.Count)
                    throw new ArchiveReviewException("Lote incompleto ou com cópias extras. Não será repetido automaticamente; originais preservados.");
                if (batch.FastCopy)
                {
                    var fastBackend = backend as IArchiveFastBatchBackend ?? throw new ArchiveReviewException("Identificação rápida indisponível.");
                    if (backend is IArchiveBulkFastBatchBackend bulk) bulk.RefreshTargetIdentities(volume, candidates);
                    var capturedFast = candidates.Select(x => (Identity: x, Hash: fastBackend.CaptureTargetIdentity(volume, rows[0], x).Hash)).ToList();
                    var usedFast = batch.Matches.Values.Select(x => x.EntryId).ToHashSet(StringComparer.Ordinal);
                    var matchesFast = new Dictionary<string, DestinationIdentity>(batch.Matches, StringComparer.Ordinal);
                    foreach (var row in rows.OrderBy(x => x.Key, StringComparer.Ordinal))
                    {
                        if (matchesFast.ContainsKey(row.Key)) continue;
                        ValidateNativeSource(job, row);
                        var candidate = capturedFast.FirstOrDefault(x => x.Hash == row.FastIdentityHash && !usedFast.Contains(x.Identity.EntryId));
                        if (candidate.Identity is null) { Review(row, "Identificação do lote difere da origem; cópias preservadas para revisão, sem repetir a transferência."); continue; }
                        usedFast.Add(candidate.Identity.EntryId); matchesFast.Add(row.Key, candidate.Identity);
                    }
                    batch.Matches = matchesFast;
                }
                else
                {
                var expected = rows.ToDictionary(x => x.Key, x => OriginalSnapshot(job, x), StringComparer.Ordinal);
                var captured = candidates.OrderBy(x => x.EntryId, StringComparer.Ordinal).Select(candidate => new
                { Identity = candidate, Snapshot = backend.TargetSnapshot(volume, rows[0], candidate, Evidence(job, rows[0], "batch-candidate"), false) }).ToList();
                var used = batch.Matches.Values.Select(x => x.EntryId).ToHashSet(StringComparer.Ordinal); var matches = new Dictionary<string, DestinationIdentity>(batch.Matches, StringComparer.Ordinal);
                var claimed = ledger.GetItems(job.Id).Where(x => x.NativeBatchId != batch.Id && x.TargetEntryId is not null).Select(x => x.TargetEntryId!).ToHashSet(StringComparer.Ordinal);
                foreach (var row in rows.OrderBy(x => x.Key, StringComparer.Ordinal))
                {
                    if (matches.ContainsKey(row.Key)) continue;
                    ValidateNativeSource(job, row);
                    var copy = captured.FirstOrDefault(x => !used.Contains(x.Identity.EntryId) && !claimed.Contains(x.Identity.EntryId) &&
                        RelevantMessageValidation.Equivalent(expected[row.Key], x.Snapshot));
                    if (copy is null) { Review(row, "Informações relevantes desta mensagem não correspondem à origem; cópia preservada para revisão."); continue; }
                    used.Add(copy.Identity.EntryId); matches.Add(row.Key, copy.Identity);
                }
                batch.Matches = matches;
                }
                batch.State = ArchiveBatchState.Matched; batch.Error = null; ledger.UpdateBatch(batch);
                Fault?.Invoke("after-batch-matched", rows[0]);
            }
            if (batch.State == ArchiveBatchState.Matched)
            {
                if (batch.Matches.Keys.Any(x => !batch.Items.Contains(x)))
                    throw new ArchiveReviewException("Mapa durável do lote incompleto.");
                Dictionary<string, DestinationIdentity>? finalized = null;
                var toFinalize = rows.Where(x => batch.Matches.ContainsKey(x.Key) && x.RepairToken is null && x.State is not (ArchiveItemState.Verified or ArchiveItemState.Removed or ArchiveItemState.InDeletedItems or ArchiveItemState.BackupVerifiedSourceUnmanaged) && !(x.TargetEntryId is not null && x.TargetFolderId!=batch.FolderId && x.TargetSnapshotHash is not null)).ToList();
                if (batch.FastCopy && backend is IArchiveBulkFinalizeBackend bulkFinal && toFinalize.Count > 0)
                {
                    finalized = bulkFinal.FinalizeBatchCopies(volume, batch, toFinalize);
                    Fault?.Invoke("after-batch-finalize-before-checkpoint", toFinalize[0]);
                }
                foreach (var row in rows)
                {
                    if (row.RepairToken is not null || row.State is ArchiveItemState.Verified or ArchiveItemState.Removed or ArchiveItemState.InDeletedItems or ArchiveItemState.BackupVerifiedSourceUnmanaged || row.TargetEntryId is not null && row.TargetFolderId!=batch.FolderId && row.TargetSnapshotHash is not null || row.State == ArchiveItemState.Copied && row.TargetEntryId is not null && batch.FastCopy) continue;
                    if (!batch.Matches.TryGetValue(row.Key, out var mapped)) continue;
                    DestinationIdentity final;
                    if (finalized is not null) final = finalized[row.Key];
                    else
                    {
                        native.MarkBatchCopy(volume, batch, row, mapped);
                        Fault?.Invoke("after-batch-mark-before-checkpoint", row);
                        final = native.FinalizeBatchCopy(volume, batch, row, mapped);
                        Fault?.Invoke("after-batch-finalize-before-checkpoint", row);
                    }
                    row.TargetStoreId = final.StoreId; row.TargetFolderId = final.FolderId; row.TargetEntryId = final.EntryId;
                    row.State = ArchiveItemState.Copied; row.Error = batch.FastCopy ? "Cópia rápida concluída; validação completa pendente antes da remoção dos originais." : null; ledger.UpdateItem(row);
                    if (!batch.FastCopy && row.TargetSnapshotHash is null)
                    {
                        var snapshot = backend.TargetSnapshot(volume, row, final, Evidence(job, row, "target-baseline"));
                        if (!RelevantMessageValidation.Equivalent(OriginalSnapshot(job, row), snapshot))
                            throw new ArchiveReviewException("Identificação ou finalização alterou informações relevantes da cópia.");
                        row.TargetSnapshotHash = snapshot.Hash;
                        row.TargetFieldsPath = Path.Combine(Path.GetDirectoryName(snapshot.MessagePath)!, "field-hashes.json");
                        SaveRelevantTarget(row, snapshot); ledger.UpdateItem(row);
                    }
                    Console.WriteLine($"Cópia registrada em lote: {row.Key[..12]}, parte {volume.Number}.");
                }
                volume.ObservedBytes = backend.MeasureVolume(volume);
                if (volume.ObservedBytes > config.HardLimitBytes - config.ReserveBytes)
                {
                    volume.State = ArchiveVolumeState.ReviewRequired; ledger.UpdateVolume(volume);
                    throw new ArchiveReviewException("PST ultrapassou o limite observado; originais preservados.");
                }
                ledger.UpdateVolume(volume);
                bool complete = batch.Matches.Count == rows.Count;
                batch.State = complete ? ArchiveBatchState.Complete : ArchiveBatchState.ReviewRequired;
                batch.Error = complete ? null : "Parte do lote exige revisão; mensagens correspondentes foram registradas e os demais lotes podem continuar.";
                ledger.UpdateBatch(batch);
                if (complete) native.CloseBatch(volume, batch);
            }
            return true;
        }
        catch (ArchiveInjectedCrashException) { throw; }
        catch (ArchiveReviewException ex)
        {
            batch.State = ArchiveBatchState.ReviewRequired; batch.Error = ex.Message; ledger.UpdateBatch(batch);
            foreach (var row in rows.Where(x => x.RepairToken is null && x.State is not (ArchiveItemState.Verified or ArchiveItemState.Removed or ArchiveItemState.InDeletedItems or ArchiveItemState.BackupVerifiedSourceUnmanaged or ArchiveItemState.SkippedByPolicy) && !(x.State == ArchiveItemState.Copied && x.TargetEntryId is not null && (batch.FastCopy || x.TargetSnapshotHash is not null)))) Review(row, ex.Message);
            return false;
        }
        catch (Exception ex)
        {
            batch.Error = "interrupted:" + ex.GetType().Name + ":" + ex.Message; ledger.UpdateBatch(batch);
            Console.Error.WriteLine("Lote interrompido; retome esta execução para conferir as cópias existentes. " + ex.GetType().Name);
            return false;
        }
        finally { Console.WriteLine($"Lote {batch.Id[..8]}: {rows.Count} itens processados em {batchClock.Elapsed.TotalSeconds:F1} s; estado {batch.State}."); }
    }
    private void ValidateBatchManifest(ArchiveJob job, ArchiveBatch batch, List<ArchiveItem> rows)
    {
        if (!Guid.TryParseExact(batch.Id, "N", out _) || batch.JobId != job.Id ||
            Path.GetFullPath(batch.DirectoryPath) != Path.Combine(job.RootDirectory, "batches", batch.Id)) throw new ArchiveReviewException("Identidade/pasta do lote inválida.");
        string manifestPath = ArchiveFiles.Child(job.RootDirectory, Path.Combine(batch.DirectoryPath, "manifest.json"));
        if (ArchiveFiles.Sha(manifestPath) != batch.ManifestSha256) throw new ArchiveReviewException("Manifesto do lote foi alterado.");
        var manifest = JsonSerializer.Deserialize<BatchManifest>(File.ReadAllText(manifestPath)) ?? throw new ArchiveReviewException("Manifesto ausente.");
        rows=rows.Select(row=>BatchOriginalEvidence(job,row)).ToList();
        if (manifest.Id != batch.Id || manifest.JobId != job.Id || manifest.VolumeNumber != batch.VolumeNumber ||
            manifest.FastCopy != batch.FastCopy || manifest.SourceStoreId != batch.SourceStoreId || manifest.SourceFolderId != batch.SourceFolderId || !manifest.RelativePath.SequenceEqual(batch.RelativePath) ||
            !manifest.Members.Select(x => x.Key).SequenceEqual(batch.Items) || rows.Any(row => row.NativeBatchId != batch.Id || row.VolumeNumber != batch.VolumeNumber || row.SourceStoreId != batch.SourceStoreId ||
                row.SourceFolderId != batch.SourceFolderId || !row.RelativePath.SequenceEqual(batch.RelativePath)))
            throw new ArchiveReviewException("Escopo do lote difere do plano congelado.");
        foreach (var member in manifest.Members)
        {
            var row = rows.Single(x => x.Key == member.Key);
            if (row.SourceEntryId != member.EntryId || (!batch.FastCopy && (row.SnapshotHash != member.SnapshotHash || row.StageMsgSha256 != member.MsgSha256)) ||
                (batch.FastCopy ? row.FastIdentityHash : row.SourceRelevantHash) != member.RelevantHash || row.ReservedBytes != member.Reservation) throw new ArchiveReviewException("Evidências/reservas do lote foram alteradas.");
            if (batch.FastCopy) CheckFastIdentity(job, row);
            else { CheckBackupFile(job, row); _ = OriginalSnapshot(job, row); }
        }
    }
    private ArchiveItem BatchOriginalEvidence(ArchiveJob job,ArchiveItem row)
    {
        if(row.RepairToken is null)return row;
        string path=ArchiveFiles.Child(job.RootDirectory,row.RepairManifestPath ?? throw new ArchiveReviewException("Histórico do reparo ausente."));
        if(ArchiveFiles.Sha(path)!=row.RepairManifestSha256)throw new ArchiveReviewException("Histórico do reparo alterado.");
        using var document=JsonDocument.Parse(File.ReadAllText(path));
        var previous=document.RootElement.GetProperty("Previous").Deserialize<ArchiveItem>() ?? throw new ArchiveReviewException("Baseline anterior ausente.");
        var source=document.RootElement.GetProperty("Source").Deserialize<ArchiveSnapshot>() ?? throw new ArchiveReviewException("Origem do reparo ausente.");
        if(previous.RepairToken is not null || previous.JobId!=row.JobId || previous.Key!=row.Key || previous.Token!=row.Token || previous.NativeBatchId!=row.NativeBatchId ||
            previous.SourceStoreId!=row.SourceStoreId || previous.SourceFolderId!=row.SourceFolderId || previous.SourceEntryId!=row.SourceEntryId || previous.SourceRecordKey!=row.SourceRecordKey ||
            document.RootElement.GetProperty("DestinationToken").GetString()!=row.RepairToken || source.Hash!=row.SnapshotHash || source.MessageSha256!=row.StageMsgSha256 ||
            Path.GetFullPath(source.MessagePath)!=Path.GetFullPath(row.StageMsgPath!))throw new ArchiveReviewException("Reparo não corresponde ao membro original do lote.");
        CheckBackupFile(job,previous);_ = OriginalSnapshot(job,previous);CheckBackupFile(job,row);_ = OriginalSnapshot(job,row);
        return previous;
    }
    private void ValidateNativeSource(ArchiveJob job, ArchiveItem row)
    {
        if (row.FastIdentityHash is not null)
        {
            CheckFastIdentity(job, row);
            if (backend is not IArchiveBulkFastBatchBackend && !string.IsNullOrEmpty(row.SourceChangeKey) && backend.SourceMatchesChangeKey(row, row.SourceChangeKey)) return;
            var fast = backend as IArchiveFastBatchBackend ?? throw new ArchiveReviewException("Identificação rápida indisponível.");
            var currentIdentity = fast.CaptureSourceIdentity(row);
            if (currentIdentity.Hash != row.FastIdentityHash)
            {
                var initial = ArchiveFiles.ReadFields(ArchiveFiles.Child(job.RootDirectory, row.FastIdentityPath!));
                var differences = initial.Keys.Union(currentIdentity.Fields.Keys)
                    .Where(key => initial.GetValueOrDefault(key) != currentIdentity.Fields.GetValueOrDefault(key));
                throw new ArchiveReviewException("Origem mudou após a identificação do lote; campos divergentes: " + string.Join(", ", differences) + ". Cópias preservadas.");
            }
            return;
        }
        CheckBackupFile(job, row);
        if (string.IsNullOrEmpty(row.SourceChangeKey) || !backend.SourceMatchesChangeKey(row, row.SourceChangeKey))
        {
            var current = backend.SourceSnapshot(row, null, Evidence(job, row, "batch-source-current"));
            if (!RelevantMessageValidation.Equivalent(OriginalSnapshot(job, row), current))
                throw new ArchiveReviewException("Origem mudou após o backup; lote não será adotado.");
        }
    }
    private static long BatchEstimate(ArchiveItem row) => row.StageMsgPath is null ? Math.Max(1, row.EstimatedBytes) : new FileInfo(row.StageMsgPath).Length;
    private void CheckFastIdentity(ArchiveJob job, ArchiveItem row)
    {
        var path = ArchiveFiles.Child(job.RootDirectory, row.FastIdentityPath ?? throw new ArchiveReviewException("Identificação rápida ausente."));
        if (ArchiveFiles.HashFields(ArchiveFiles.ReadFields(path)) != row.FastIdentityHash)
            throw new ArchiveReviewException("Identificação rápida foi alterada.");
    }
    private void HydrateFastEvidence(ArchiveJob job, ArchiveItem row, ArchiveVolume volume)
    {
        ValidateNativeSource(job, row);
        if (row.StageMsgPath is null)
        {
            var original = backend.Export(row, Evidence(job, row, "original-deferred"));
            row.StageMsgPath = ArchiveFiles.Child(job.RootDirectory, original.MessagePath);
            row.StageMsgSha256 = original.MessageSha256; row.SnapshotHash = original.Hash;
            row.SourceChangeKey = original.ChangeKey; SaveRelevantOriginal(row, original); ledger.UpdateItem(row);
        }
        if (row.TargetSnapshotHash is null)
        {
            var target = backend.TargetSnapshot(volume, row, Unique(volume, row), Evidence(job, row, "target-deferred"));
            var comparable = ComparableOriginal(job, row);
            if (!RelevantMessageValidation.Equivalent(comparable, target))
                throw new ArchiveReviewException("Validação completa encontrou diferenças entre origem e PST; diferenças: " + RelevantMessageValidation.DifferenceSummary(comparable, target) + ". Originais preservados.");
            row.TargetSnapshotHash = target.Hash;
            row.TargetFieldsPath = Path.Combine(Path.GetDirectoryName(target.MessagePath)!, "field-hashes.json");
            SaveRelevantTarget(row, target); ledger.UpdateItem(row);
        }
    }
    private bool RecordedBackupMatches(ArchiveJob job, ArchiveItem row, ArchiveSnapshot backup)
    {
        if (backup.MessageSha256 != row.StageMsgSha256) return false;
        if (backup.Hash == row.SnapshotHash) return true;
        return EmbeddedMessageEvidence.CanReplay(OriginalSnapshot(job, row), backup);
    }
    private ArchiveSnapshot ComparableOriginal(ArchiveJob job, ArchiveItem row)
    {
        var original = OriginalSnapshot(job, row);
        if (EmbeddedMessageEvidence.Indices(original.Fields).Count == 0 || EmbeddedMessageEvidence.Complete(original.Fields, original.RelevantFields)) return original;
        var replay = backend.BackupSnapshot(row, Evidence(job, row, "embedded-original-replay"));
        if (!RecordedBackupMatches(job, row, replay) || !EmbeddedMessageEvidence.CanReplay(original, replay))
            throw new ArchiveReviewException("Evidência recursiva do backup original não foi confirmada.");
        return original with { RelevantFields = replay.RelevantFields };
    }
    private ArchiveSnapshot OriginalSnapshot(ArchiveJob job, ArchiveItem row)
    {
        var fields = Original(job, row); Dictionary<string, string>? relevant = null;
        if (row.SourceRelevantHash is not null)
        {
            string sidecar = ArchiveFiles.Child(job.RootDirectory, Path.Combine(Path.GetDirectoryName(row.StageMsgPath!)!, "relevant-fields.json"));
            relevant = ArchiveFiles.ReadFields(sidecar);
            if (ArchiveFiles.HashFields(relevant) != row.SourceRelevantHash) throw new ArchiveReviewException("Evidência das informações relevantes da origem foi alterada.");
        }
        var original=DeliveryReportEvidence.BindSubjectProof(new(fields, row.StageMsgPath!, row.StageMsgSha256!, row.SourceChangeKey, relevant));
        return row.RepairToken is not null && backend is OutlookArchiveBackend ? RelevantMessageValidation.RecoveryProfile(original) : original;
    }
    private static void SaveRelevantOriginal(ArchiveItem row, ArchiveSnapshot snapshot)
    {
        if (snapshot.RelevantFields is null) return;
        SaveRelevantSidecar(Path.Combine(Path.GetDirectoryName(row.StageMsgPath!)!, "relevant-fields.json"), snapshot.RelevantFields);
        row.SourceRelevantHash = ArchiveFiles.HashFields(snapshot.RelevantFields);
    }
    private static void SaveRelevantTarget(ArchiveItem row, ArchiveSnapshot snapshot)
    {
        if (snapshot.RelevantFields is null) return;
        SaveRelevantSidecar(Path.Combine(Path.GetDirectoryName(row.TargetFieldsPath!)!, "relevant-fields.json"), snapshot.RelevantFields);
        row.TargetRelevantHash = ArchiveFiles.HashFields(snapshot.RelevantFields);
    }
    private Dictionary<string, string>? TargetRelevantBaseline(ArchiveJob job, ArchiveItem row)
    {
        if (row.TargetRelevantHash is null) return null;
        string path = ArchiveFiles.Child(job.RootDirectory, Path.Combine(Path.GetDirectoryName(row.TargetFieldsPath!)!, "relevant-fields.json"));
        var fields = ArchiveFiles.ReadFields(path);
        if (ArchiveFiles.HashFields(fields) != row.TargetRelevantHash) throw new ArchiveReviewException("Evidência das informações relevantes do destino foi alterada.");
        string message=Path.Combine(Path.GetDirectoryName(path)!,"message.msg");
        if(row.RepairToken is not null && backend is OutlookArchiveBackend && File.Exists(message))
            return RelevantMessageValidation.RecoveryProfile(new(TargetBaseline(job,row),message,ArchiveFiles.Sha(message),null,fields)).RelevantFields;
        return fields;
    }
    private static void SaveRelevantSidecar(string path, Dictionary<string, string> fields)
    {
        if (!File.Exists(path)) ArchiveFiles.WriteNew(path, fields);
        else if (ArchiveFiles.HashFields(ArchiveFiles.ReadFields(path)) != ArchiveFiles.HashFields(fields))
            throw new ArchiveReviewException("Evidência relevante já existe com conteúdo diferente.");
    }
}
