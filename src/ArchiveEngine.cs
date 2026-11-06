using System.Text.Json;

internal sealed partial class ArchiveEngine : IDisposable
{
    private readonly string root;
    private readonly FileStream lease;
    private readonly ArchiveLedger ledger;
    private readonly IArchiveBackend backend;
    private sealed record ReconciliationCandidate(DestinationIdentity Identity, ArchiveSnapshot Snapshot);
    private Dictionary<string, List<ReconciliationCandidate>>? reconciliationCandidates;
    // Used by tests to stop exactly between COM and the SQLite checkpoint.
    internal Action<string, ArchiveItem>? Fault;
    internal ArchiveEngine(string rootDirectory, IArchiveBackend backend)
    {
        ArchiveFiles.ValidateLocalDirectory(rootDirectory); root = Path.GetFullPath(rootDirectory).TrimEnd(Path.DirectorySeparatorChar);
        Directory.CreateDirectory(root);
        lease = new FileStream(Path.Combine(root, "archive.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
        try { ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")); this.backend = backend; }
        catch { lease.Dispose(); throw; }
    }
    internal string Plan(ArchiveConfig config)
    {
        config.Validate();
        if (!string.Equals(Path.GetFullPath(config.RootDirectory), root, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Raiz da configuração não corresponde ao catálogo aberto.");
        return BeginPlanning(config);
    }
    private string CreatePlan(ArchiveConfig config, List<FolderScopeSnapshot> initialScope, IReadOnlyList<FolderScopeSnapshot> selectedScope)
    {
        return BeginPlanning(config,initialScope,selectedScope);
    }
    private List<FolderScopeSnapshot> InitialScope(ArchiveJob job)
    {
        var envelope = Envelope(job);
        string path = Path.Combine(job.RootDirectory, envelope.InitialScopeSha256 is null ? "source-plan.json" : "source-scope.json");
        if (ArchiveFiles.Sha(path) != (envelope.InitialScopeSha256 ?? envelope.PlanSha256))
            throw new ArchiveReviewException("Evidência do escopo inicial foi alterada.");
        if (envelope.InitialScopeSha256 is not null)
            return JsonSerializer.Deserialize<List<FolderScopeSnapshot>>(File.ReadAllText(path)) ?? throw new ArchiveReviewException("Escopo inicial inválido.");
        // Backward-compatible with plans created before the complete identity snapshot.
        return JsonSerializer.Deserialize<List<SourceDescriptor>>(File.ReadAllText(path))!
            .GroupBy(x => x.FolderId).Select(g => new FolderScopeSnapshot(g.Key, g.First().RelativePath, g.Select(x => x.EntryId).ToList())).ToList();
    }
    private List<FolderScopeSnapshot> NewScope(ArchiveJob job, List<FolderScopeSnapshot> current)
    {
        var config = Envelope(job).Config;
        if (ledger.HasUnreconciledRemoval(config.SourceStoreId))
            throw new ArchiveReviewException("Há uma remoção interrompida sem identidade reconciliada nesta origem. Retome ou revise essa execução antes de preparar outro plano.");
        var original = InitialScope(job).SelectMany(x => x.EntryIds).ToHashSet(StringComparer.Ordinal);
        var registered = ledger.GetKnownItemKeys();
        return current.Select(folder => new FolderScopeSnapshot(folder.FolderId, folder.RelativePath,
            folder.EntryIds.Where(id => !original.Contains(id) && !registered.Contains(ArchiveLedger.ComputeItemKey(config.SourceStoreId, id))).ToList()))
            .Where(x => x.EntryIds.Count > 0).ToList();
    }
    internal ArrivalCheckResult CheckNew(string jobId)
    {
        ArrivalCheckResult result;
        var job = ledger.GetJob(jobId);
        try
        {
            var config = Envelope(job).Config;
            var scope = NewScope(job, backend.CaptureScope(config));
            int count = backend.DiscoverFrozen(config, scope).Count(x => x.IsSupported && x.IsSent);
            result = new(count, DateTimeOffset.UtcNow);
        }
        catch (Exception ex) { result = new(0, DateTimeOffset.UtcNow, "Não foi possível conferir novos e-mails: " + ex.Message); }
        // Arrival checks are evidence, never a change to the frozen plan or item states.
        try { ArchiveFiles.WriteNew(Path.Combine(job.RootDirectory, "arrival-check-" + Guid.NewGuid().ToString("N") + ".json"), result); }
        catch (IOException) { /* Notification remains available even when a diagnostic cannot be saved. */ }
        catch (UnauthorizedAccessException) { }
        return result;
    }
    internal string PlanNew(string previousJobId)
    {
        var previous = ledger.GetJob(previousJobId); var config = Envelope(previous).Config;
        config.MoveValidatedToDeletedItems=ledger.WorkflowManifest(previousJobId)?.Move ?? config.MoveValidatedToDeletedItems;
        // A follow-up plan intentionally reuses the destination created by
        // the previous plan, while preserving its independent evidence.
        if(config.DestinationPstPath is { } destination && File.Exists(destination))config.ExistingPstPath=destination;
        var current = backend.CaptureScope(config);
        return CreatePlan(config, current, NewScope(previous, current));
    }
    private ArchivePlanEnvelope Envelope(ArchiveJob job,bool allowPlanning=false)
    {
        if(job.PlanningPending && !allowPlanning)throw new ArchiveReviewException("O planejamento está incompleto. Use Continuar arquivamento para concluir a preparação antes de copiar ou remover mensagens.");
        var envelope = JsonSerializer.Deserialize<ArchivePlanEnvelope>(job.ConfigJson) ?? throw new InvalidDataException("Configuração ausente.");
        envelope.Config.Validate();
        if (!string.Equals(Path.GetFullPath(envelope.Config.RootDirectory).TrimEnd(Path.DirectorySeparatorChar), root, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(Path.GetFullPath(job.RootDirectory), Path.Combine(root, "jobs", job.Id), StringComparison.OrdinalIgnoreCase) || !Guid.TryParseExact(job.Id, "N", out _))
            throw new ArchiveReviewException("Raiz/identidade da execução não conferem.");
        return envelope;
    }
    private void SealPlan(ArchiveJob job)
    {
        var envelope = Envelope(job,true);
        string path = Path.Combine(job.RootDirectory, "source-plan.json");
        if (ArchiveFiles.Sha(path) != envelope.PlanSha256) throw new ArchiveReviewException("Plano congelado foi alterado.");
        var rows = JsonSerializer.Deserialize<List<SourceDescriptor>>(File.ReadAllText(path)) ?? throw new InvalidDataException("Plano inválido.");
        int registered = 0;
        var known = ledger.GetKnownItemKeys();
        foreach (var source in rows)
        {
            if (source.StoreId != envelope.Config.SourceStoreId || source.RelativePath.Length == 0 || source.RelativePath.Any(string.IsNullOrWhiteSpace))
                throw new ArchiveReviewException("Item fora da origem planejada.");
            if (source.IsSupported && source.IsSent && !ArchiveDateFilter.Includes(source.MessageDateUtc, envelope.Config.CutoffDateInclusive, envelope.Config.CutoffTimeZoneId))
                throw new ArchiveReviewException("Item fora da data de corte do plano congelado.");
            string key = ArchiveLedger.ComputeItemKey(source.StoreId, source.EntryId);
            if (!known.Contains(key))
            {
                ledger.AddItem(new ArchiveItem { JobId = job.Id, SourceStoreId = source.StoreId, SourceFolderId = source.FolderId,
                    SourceEntryId = source.EntryId, SourceRecordKey = source.RecordKey, SourceMessageClass = source.MessageClass, EstimatedBytes = source.Size, RelativePath = source.RelativePath,
                    State = DeliveryReportEvidence.IgnoredMeeting(source.MessageClass) ? ArchiveItemState.SkippedByPolicy : !source.IsSupported ? ArchiveItemState.SkippedUnsupported : !source.IsSent ? ArchiveItemState.SkippedDraft : ArchiveItemState.Planned });
                known.Add(key);
            }
            registered++;
            if (registered % 100 == 0 || registered == rows.Count)
                Console.WriteLine($"Planejamento: catálogo — {registered:N0}/{rows.Count:N0} identidades registradas ou conferidas.");
        }
    }
    internal void Copy(string jobId,bool moveInline=false,bool allowRetainedBackup=false)
    {
        if(workflowSelection is null && ledger.WorkflowManifest(jobId) is not null)throw new ArchiveReviewException("Esta execução usa lotes. Use Continuar arquivamento para manter seus checkpoints.");
        var job = ledger.GetJob(jobId); var config = Envelope(job).Config; if(workflowSelection is null)SealPlan(job); UpgradeDeliveryReports(job,config);
        if(allowRetainedBackup)config.AllowRetainedMsgBackupForRemoval=true;
        if (backend is IArchiveBatchBackend native && !CopyNativeBatches(job, config, native)) { Status(jobId); return; }
        foreach (var row in WorkItems(jobId))
        {
            using var itemTiming=ArchiveOperationTiming.Measure("item.copy",row.Key[..12]);
            if (row.NativeBatchId is not null) continue;
            if (row.State is not (ArchiveItemState.Planned or ArchiveItemState.Exported or ArchiveItemState.ImportStarted or ArchiveItemState.Copied)) continue;
            // Only this invocation can know that Import has not been attempted.
            // A persisted ImportStarted on resume remains ambiguous even when
            // a lookup finds no token, so it must retain recovery safeguards.
            bool freshImportIntent = false;
            ArchiveSnapshot? freshSource=null,freshTarget=null;
            try
            {
                if (row.State == ArchiveItemState.Planned)
                {
                    CheckSpace(row.EstimatedBytes, config);
                    var original = backend.Export(row, Evidence(job, row, "original"));
                    freshSource=original;
                    row.StageMsgPath = ArchiveFiles.Child(job.RootDirectory, original.MessagePath);
                    row.StageMsgSha256 = original.MessageSha256; row.SnapshotHash = original.Hash;
                    row.SourceChangeKey = original.ChangeKey;
                    SaveRelevantOriginal(row, original);
                    row.State = ArchiveItemState.Exported; row.Error = null; ledger.UpdateItem(row); Fault?.Invoke("after-export", row);
                }
                CheckBackupFile(job, row);
                if (row.State == ArchiveItemState.Exported)
                {
                    // Re-exporting unchanged Exchange messages repeats SaveAs,
                    // attachment downloads and hashing. A matching provider
                    // ChangeKey proves no intervening change. Missing/different
                    // keys retain the full comparison; final Verify/Remove always
                    // compare complete snapshots regardless of this shortcut.
                    if (string.IsNullOrEmpty(row.SourceChangeKey) || !backend.SourceMatchesChangeKey(row, row.SourceChangeKey))
                    {
                        var current = backend.SourceSnapshot(row, null, Evidence(job, row, "source-before-import"));
                        if (!RelevantMessageValidation.Equivalent(OriginalSnapshot(job, row), current)) throw new ArchiveReviewException("Origem mudou depois da exportação; importação bloqueada.");
                    }
                    var volume = Allocate(job, row, config);
                    row.VolumeNumber = volume.Number; row.State = ArchiveItemState.ImportStarted; row.Error = null;
                    ledger.UpdateItem(row); freshImportIntent = true; Fault?.Invoke("after-import-intent", row);
                }
                var targetVolume = Volume(row);
                backend.AttachVolume(targetVolume); ledger.UpdateVolume(targetVolume);
                if (row.State == ArchiveItemState.ImportStarted)
                {
                    var existing = backend.FindToken(targetVolume, row);
                    if (existing.Matches.Count > 1)
                        throw new ArchiveReviewException("Importação ambígua: token repetido ou mensagem sem identificação no destino.");
                    var identity = existing.Matches.SingleOrDefault() ??
                        (!freshImportIntent && existing.UnmarkedItems > 0 ? RecoverUnmarked(job, row, targetVolume) : backend.Import(targetVolume, row));
                    Fault?.Invoke("after-import-before-checkpoint", row);
                    row.TargetEntryId = identity.EntryId; row.TargetFolderId = identity.FolderId; row.TargetStoreId = identity.StoreId;
                    row.State = ArchiveItemState.Copied; ledger.UpdateItem(row);
                }
                var found = Unique(targetVolume, row);
                if (row.TargetSnapshotHash is null)
                {
                    var copy = backend.TargetSnapshot(targetVolume, row, found, Evidence(job, row, "target-baseline"));
                    freshTarget=copy;
                    row.TargetSnapshotHash = copy.Hash;
                    row.TargetFieldsPath = Path.Combine(Path.GetDirectoryName(copy.MessagePath)!, "field-hashes.json");
                    SaveRelevantTarget(row, copy);
                    row.Error = null; ledger.UpdateItem(row);
                }
                targetVolume.ObservedBytes = backend.MeasureVolume(targetVolume);
                if (targetVolume.ObservedBytes > config.HardLimitBytes - config.ReserveBytes)
                {
                    targetVolume.State = ArchiveVolumeState.ReviewRequired; ledger.UpdateVolume(targetVolume);
                    throw new ArchiveReviewException("PST ultrapassou o limite observado; não haverá remoção de originais.");
                }
                ledger.UpdateVolume(targetVolume);
                Console.WriteLine($"Cópia registrada: {row.Key[..12]}, parte {targetVolume.Number}.");
                if(moveInline && config.AllowRetainedMsgBackupForRemoval && freshSource is not null && freshTarget is not null)
                    VerifyFreshCopyAndMove(job,row,targetVolume,config,freshSource,freshTarget);
                workflowProgress?.Invoke();
            }
            catch (ArchiveInjectedCrashException) { throw; }
            catch (ArchiveReviewException ex) { Review(row, ex.Message); }
            catch (Exception ex)
            {
                row.Error = "interrupted:" + ex.GetType().Name + ":" + ex.Message; ledger.UpdateItem(row);
                Console.Error.WriteLine("Interrupção registrada. Execute novamente a mesma execução para reconciliar: " + ex.GetType().Name);
                break;
            }
        }
        Status(jobId);
    }
    private ArchiveVolume Allocate(ArchiveJob job, ArchiveItem row, ArchiveConfig config)
    {
        long estimate = Math.Max(row.EstimatedBytes, row.StageMsgPath is null ? 0 : new FileInfo(row.StageMsgPath).Length);
        var volumes = ledger.GetVolumes(job.Id);
        var volume = row.VolumeNumber.HasValue ? volumes.Single(x => x.Number == row.VolumeNumber) : volumes.LastOrDefault();
        if (volume is null) volume = NewVolume(job, 1);
        EnsureVolume(volume);
        // Outlook can defer physical file growth. Never discard reservations
        // merely because an import returned: a small cached file is not free space.
        long reserved = ledger.GetReservedBytes(job.Id, volume.Number, row.Key);
        bool empty = reserved == 0 && !volume.IsExisting;
        long size = backend.MeasureVolume(volume);
        long budget = Math.Max(size, checked(volume.InitialBytes + reserved));
        var decision = ArchiveVolumePolicy.Decide(budget, 0, estimate, config.TargetVolumeBytes, config.HardLimitBytes, config.ReserveBytes,
            empty ? Math.Max(size, ArchiveVolumePolicy.DefaultEmptyVolumeBaseBytes) : ArchiveVolumePolicy.DefaultEmptyVolumeBaseBytes);
        bool rollExisting = volume.IsExisting && budget >= config.TargetVolumeBytes;
        if (!decision.Allowed && !rollExisting) throw new ArchiveReviewException("Item não cabe no limite seguro: " + decision.Reason);
        if (decision.NewVolume || rollExisting)
        {
            volume = NewVolume(job, volumes.Count == 0 ? 2 : volumes.Max(x => x.Number) + 1); EnsureVolume(volume);
            size = backend.MeasureVolume(volume);
            decision = ArchiveVolumePolicy.Decide(size, 0, estimate, config.TargetVolumeBytes, config.HardLimitBytes, config.ReserveBytes, size);
            if (!decision.Allowed || decision.NewVolume) throw new ArchiveReviewException("Parte nova não comporta o item com segurança.");
        }
        row.ReservedBytes = decision.ReservationBytes;
        return volume;
    }
    private DestinationIdentity RecoverUnmarked(ArchiveJob job, ArchiveItem row, ArchiveVolume volume)
    {
        CheckBackupFile(job, row); Original(job, row);
        var source = backend.SourceSnapshot(row, null, Evidence(job, row, "source-reconciliation"));
        // Old plans can lack the optional metadata profile. Re-read the same
        // SHA-verified immutable MSG for a current profile instead of treating
        // its absent sidecar as a changed source. The original field baseline,
        // backup bytes and complete source/backup equality are still required.
        var expected = backend.BackupSnapshot(row, Evidence(job, row, "backup-reconciliation"));
        if (!RecordedBackupMatches(job, row, expected) || !RelevantMessageValidation.Equivalent(expected, source))
            throw new ArchiveReviewException("Origem e backup não correspondem durante a reconciliação.");
        var claimed = ledger.GetItems(job.Id).Where(x => x.Key != row.Key && x.TargetEntryId is not null).Select(x => x.TargetEntryId).ToHashSet(StringComparer.Ordinal);
        string group = volume.PstPath + "|" + JsonSerializer.Serialize(row.RelativePath);
        List<ReconciliationCandidate>? captured = reconciliationCandidates?.GetValueOrDefault(group);
        if (captured is null)
        {
            captured = [];
            var identities = backend.UnmarkedCandidates(volume, row).Where(x => !claimed.Contains(x.EntryId)).ToList();
            for (int i = 0; i < identities.Count; i++)
            {
                var candidate = identities[i];
                Console.WriteLine($"Conferindo cópias antigas sem identificador: {i + 1}/{identities.Count} nesta pasta…");
                var snapshot = backend.TargetSnapshot(volume, row, candidate, Evidence(job, row, "unmarked-candidate"), false);
                captured.Add(new(candidate, snapshot));
            }
            if (reconciliationCandidates is not null) reconciliationCandidates[group] = captured;
        }
        var matches = captured.Where(x => !claimed.Contains(x.Identity.EntryId) && RelevantMessageValidation.Equivalent(expected, x.Snapshot)).Select(x => x.Identity).ToList();
        if (matches.Count == 0) throw new ArchiveReviewException("Nenhuma cópia sem identificador corresponde integralmente a este backup no destino previsto; o backup MSG e o original foram preservados. Não será feita uma nova cópia automaticamente.");
        if (matches.Count > 1) throw new ArchiveReviewException($"Há {matches.Count} cópias correspondentes sem identificação única; a revisão é necessária para evitar duplicatas.");
        backend.MarkCopy(volume, row, matches[0]);
        var marked = backend.TargetSnapshot(volume, row, matches[0], Evidence(job, row, "reidentified-copy"));
        if (!RelevantMessageValidation.Equivalent(expected, marked)) throw new ArchiveReviewException("Gravação do identificador alterou informações relevantes da cópia; revisão obrigatória.");
        return matches[0];
    }
    internal void Reconcile(string jobId, IReadOnlySet<string>? selectedKeys = null)
    {
        var bulk = backend as IArchiveBulkFastBatchBackend;
        bulk?.ClearFastIdentityCache();
        reconciliationCandidates = new(StringComparer.Ordinal);
        try { ReconcileCore(jobId, selectedKeys); }
        finally { reconciliationCandidates = null; bulk?.ClearFastIdentityCache(); }
    }
    private void ReconcileCore(string jobId, IReadOnlySet<string>? selectedKeys)
    {
        if(backend is OutlookArchiveBackend lookup)lookup.BeginSourceLookupBatch();
        var job = ledger.GetJob(jobId); Envelope(job);
        if (backend is IArchiveBatchBackend native)
        {
            foreach (var batch in ledger.GetBatches(jobId).Where(x => x.State == ArchiveBatchState.ReviewRequired && x.FolderId is not null && (selectedKeys is null || x.Items.All(selectedKeys.Contains))))
            {
                // Explicit reconciliation may adopt an existing complete result,
                // but never sends a CopyIntent to the native copier again.
                batch.State = batch.Matches.Count == batch.Items.Count ? ArchiveBatchState.Matched : ArchiveBatchState.CopyIntent;
                batch.Error = null; ledger.UpdateBatch(batch);
                if (!RunNativeBatch(job, Envelope(job).Config, native, batch) && batch.State != ArchiveBatchState.ReviewRequired) return;
            }
        }
        var reviewRows = ledger.GetItems(jobId).Where(x => x.State == ArchiveItemState.ReviewRequired && (selectedKeys is null || selectedKeys.Contains(x.Key)) && x.VolumeNumber is not null && (x.StageMsgPath is not null || x.FastIdentityHash is not null)).ToList();
        if (backend is IArchiveBulkFastBatchBackend bulkIdentities)
        {
            // Match the original native FILETIME precision. OOM DATE values
            // round to milliseconds and cannot be compared to the frozen hash.
            var deferred = reviewRows.Where(x => x.StageMsgPath is null && x.FastIdentityHash is not null).ToList();
            if (deferred.Count > 0) bulkIdentities.RefreshSourceIdentities(deferred);
        }
        var checkedPartialBatches=new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in reviewRows)
        {
            if (row.NativeBatchId is not null)
            {
                var batch = ledger.GetBatches(jobId).SingleOrDefault(x => x.Id == row.NativeBatchId);
                // A completed transfer may later fail deep verification. Recheck
                // its unique, recorded destination; never send it to the copier.
                if (batch is null || row.TargetEntryId is null) continue;
                if(batch.State!=ArchiveBatchState.Complete)
                {
                    if(row.TargetSnapshotHash is null || row.TargetFieldsPath is null || row.TargetFolderId==batch.FolderId)continue;
                    if(checkedPartialBatches.Add(batch.Id))
                    {
                        ((IArchiveBatchBackend)backend).AssertBatchIdle(batch);
                        ValidateBatchManifest(job,batch,batch.Items.Select(key=>ledger.GetItem(key) ?? throw new ArchiveReviewException("Membro do lote ausente.")).ToList());
                    }
                }
            }
            Console.WriteLine($"Reconciliando mensagem {row.Key[..12]}: conferindo origem, backup e cópia existentes…");
            // Explicit command, never reset changed source/baseline evidence.
            // Ownership + complete field equality + unique unclaimed candidate
            // are required; an email subject/date is never a deduplication key.
            string phase = "conferência do backup MSG";
            try
            {
                var volume = Volume(row); EnsureVolume(volume);
                if (row.StageMsgPath is null && row.FastIdentityHash is not null)
                {
                    phase = "complementação da evidência da cópia existente";
                    HydrateFastEvidence(job, row, volume);
                }
                CheckBackupFile(job, row);
                phase = "leitura da mensagem de origem";
                var location = backend.LocateSource(row);
                var source = backend.SourceSnapshot(row, location, Evidence(job, row, "reconcile-source-current"));
                phase = "leitura do backup MSG";
                var backup = backend.BackupSnapshot(row, Evidence(job, row, "reconcile-backup"));
                if (!RecordedBackupMatches(job, row, backup))
                    throw new ArchiveReviewException("O backup MSG não corresponde à evidência original; nenhuma cópia será refeita.");
                if (!RelevantMessageValidation.Equivalent(backup, source))
                    throw new ArchiveReviewException("A origem mudou após o backup; preserve a cópia e o original para revisão.");
                phase = "localização da cópia no PST";
                var found = backend.FindToken(volume, row);
                if (found.Matches.Count > 1) throw new ArchiveReviewException("Identificador repetido no PST; cópias preservadas para revisão.");
                var identity = found.Matches.SingleOrDefault() ?? RecoverUnmarked(job, row, volume);
                phase = "conferência da cópia no PST";
                var target = backend.TargetSnapshot(volume, row, identity, Evidence(job, row, "reconcile-target-current"));
                if (!RelevantMessageValidation.Equivalent(backup, target))
                    throw new ArchiveReviewException("A cópia no PST ainda difere do backup original; diferenças: " + RelevantMessageValidation.DifferenceSummary(backup, target) + ". A divergência exige revisão de conteúdo.");
                row.TargetEntryId = identity.EntryId; row.TargetStoreId = identity.StoreId; row.TargetFolderId = identity.FolderId;
                if (row.TargetSnapshotHash is null)
                { row.TargetSnapshotHash = target.Hash; row.TargetFieldsPath = Path.Combine(Path.GetDirectoryName(target.MessagePath)!, "field-hashes.json"); SaveRelevantTarget(row, target); }
                else TargetBaseline(job, row);
                bool trackedTrash=location?.InDeletedItems==true && row.RemovalEntryId==location.EntryId && row.RemovalFolderId==location.FolderId && row.RemovalRecordKey==location.RecordKey && row.RemovalProofPath is not null;
                if(trackedTrash)
                {
                    string proofPath=ArchiveFiles.Child(job.RootDirectory,row.RemovalProofPath!);
                    if(ArchiveFiles.Sha(proofPath)!=row.RemovalProofSha256)throw new ArchiveReviewException("Prova de movimentação anterior alterada.");
                    var removal=JsonSerializer.Deserialize<RemovalProof>(File.ReadAllText(proofPath)) ?? throw new ArchiveReviewException("Prova de movimentação anterior ausente.");
                    if(removal.Key!=row.Key || removal.StageSha!=row.StageMsgSha256 || (removal.SourceBaselineHash ?? removal.SourceHash)!=row.SnapshotHash || (removal.TargetBaselineHash ?? removal.TargetHash)!=row.TargetSnapshotHash)throw new ArchiveReviewException("Prova de movimentação difere da execução.");
                }
                row.State = trackedTrash ? ArchiveItemState.InDeletedItems : ArchiveItemState.Copied; row.Error = null; ledger.UpdateItem(row);
            }
            catch (Exception ex)
            {
                string reason = ex is System.Runtime.InteropServices.COMException com && com.HResult == unchecked((int)0x8004010F) && phase == "leitura da mensagem de origem"
                    ? "A mensagem de origem não foi localizada pelo identificador do plano; ela pode ter sido movida ou excluída fora da ferramenta. O backup MSG íntegro foi preservado."
                    : ex.Message;
                row.Error = "reconciliation-pending:" + phase + ": " + reason; ledger.UpdateItem(row);
                Console.Error.WriteLine($"Revisão pendente: {row.Key[..12]} {phase}: {reason}");
            }
        }
        Status(jobId);
    }
    private ArchiveVolume NewVolume(ArchiveJob job, int number)
    {
        var config = Envelope(job).Config;
        bool existing = number == 1 && config.ExistingPstPath is not null;
        string path = existing ? Path.GetFullPath(config.ExistingPstPath!) : Path.Combine(job.RootDirectory, "psts", ArchiveVolumePolicy.GetFileName(job.Id, number));
        if(config.DestinationPstPath is { } selected)
            path=number==1 ? Path.GetFullPath(selected) : Path.Combine(Path.GetDirectoryName(Path.GetFullPath(selected))!,Path.GetFileNameWithoutExtension(selected)+"-"+job.Id[..8]+"-parte-"+number.ToString("D4")+".pst");
        if (!existing && File.Exists(path)) throw new ArchiveReviewException("Arquivo de parte já existe sem registro correspondente.");
        var volume = new ArchiveVolume { JobId = job.Id, Number = number, PstPath = path, IsExisting = existing, SourceStoreId = config.SourceStoreId, State = ArchiveVolumeState.CreateStarted,
            PreserveFolderStructure = config.PreserveFolderStructure, FolderMappings = config.FolderMappings };
        if (!ledger.AddVolume(volume)) throw new ArchiveReviewException("Número de parte já registrado."); return volume;
    }
    private void EnsureVolume(ArchiveVolume volume)
    {
        if (volume.State == ArchiveVolumeState.ReviewRequired) throw new ArchiveReviewException("Parte pendente de revisão.");
        bool created = volume.State == ArchiveVolumeState.CreateStarted;
        if (created) backend.CreateVolume(volume); else backend.AttachVolume(volume);
        volume.ObservedBytes = backend.MeasureVolume(volume);
        if (created) volume.InitialBytes = volume.ObservedBytes;
        volume.State = ArchiveVolumeState.Ready; ledger.UpdateVolume(volume);
    }
    internal void Verify(string jobId, bool includeFast = true, bool allowRetainedBackup = false, int maxItems = int.MaxValue, bool pendingOnly = false, bool? moveValidatedToDeletedItems = null)
    {
        if(backend is OutlookArchiveBackend lookup)lookup.BeginSourceLookupBatch();
        (backend as IArchiveBulkFastBatchBackend)?.ClearFastIdentityCache();
        var job = ledger.GetJob(jobId); var config = Envelope(job).Config;
        if (allowRetainedBackup) config.AllowRetainedMsgBackupForRemoval = true;
        bool autoMove=moveValidatedToDeletedItems ?? config.MoveValidatedToDeletedItems;
        if(autoMove)Console.WriteLine("Validação com movimento automático para Itens Excluídos: somente originais aprovados, sem exclusão definitiva.");
        int checkedItems = 0;
        int consecutiveAutomaticFailures=0;
        if (!NativeWorkersIdle(jobId)) return;
        foreach (var volume in ledger.GetVolumes(jobId))
        {
            if(autoMove && consecutiveAutomaticFailures>=3)break;
            if (volume.State != ArchiveVolumeState.Ready) continue;
            var volumeRows = WorkItems(jobId).Where(x => x.VolumeNumber == volume.Number).ToList();
            foreach (var unsupported in volumeRows.Where(x => x.State == ArchiveItemState.Copied &&
                (x.Error?.StartsWith("verification-pending:InvalidOperationException:Anexo especial não suportado", StringComparison.Ordinal) == true ||
                 x.Error?.StartsWith("verification-pending:InvalidDataException:MSG sem stream de corpo reconhecido", StringComparison.Ordinal) == true)))
                Review(unsupported, "Formato de corpo ou anexo não suportado; revisão necessária. A verificação automática não será repetida.");
            var verificationRows = volumeRows.Where(x => x.State == ArchiveItemState.Copied || (!pendingOnly || autoMove) && x.State == ArchiveItemState.Verified || autoMove && x.State==ArchiveItemState.DeleteStarted).ToList();
            if (pendingOnly) Console.WriteLine(autoMove ? $"Validação incremental e movimento: {verificationRows.Count:N0} itens; mensagens já em Itens Excluídos serão puladas." : $"Validação incremental: {verificationRows.Count:N0} cópias pendentes; {volumeRows.Count(x => x.State == ArchiveItemState.Verified):N0} já validadas serão conferidas apenas na remoção.");
            if (verificationRows.Count == 0) continue;
            DiskRoundtrip disk;
            try
            {
                Console.WriteLine($"Validação: reabrindo e conferindo a gravação da parte {volume.Number}…");
                disk = backend.ReopenVolume(volume); volume.ObservedBytes = backend.MeasureVolume(volume);
                if (volume.ObservedBytes > config.HardLimitBytes - config.ReserveBytes) volume.State = ArchiveVolumeState.ReviewRequired;
                ledger.UpdateVolume(volume);
                if (volume.State != ArchiveVolumeState.Ready) continue;
            }
            catch (Exception ex)
            {
                foreach (var pending in verificationRows)
                { pending.Error = "volume-reopen-pending:" + ex.Message; ledger.UpdateItem(pending); }
                Console.Error.WriteLine("Validação da parte pendente: " + ex.Message); continue;
            }
            if (!disk.ExclusiveDiskRead && !config.AllowRetainedMsgBackupForRemoval)
            {
                const string reason = "destination_durable_reopen_or_independent_backup_missing";
                foreach (var pending in verificationRows)
                { pending.Error = reason; ledger.UpdateItem(pending); }
                Console.WriteLine("Validação bloqueada: o Outlook mantém o PST ocupado. Feche mensagens deste PST e tente novamente ou autorize explicitamente a prova adicional com backup MSG retido na confirmação de remoção.");
                continue;
            }
            if (maxItems != int.MaxValue) verificationRows = verificationRows.OrderByDescending(x => x.FastIdentityHash is not null).Take(maxItems - checkedItems).ToList();
            if (includeFast && backend is IArchiveBulkFastBatchBackend bulk)
            {
                var fastRows = verificationRows.Where(x => x.FastIdentityHash is not null && x.State!=ArchiveItemState.DeleteStarted && (x.StageMsgPath is null || x.TargetSnapshotHash is null)).ToList();
                bulk.RefreshSourceIdentities(fastRows);
            }
            foreach (var row in verificationRows)
            {
                if (checkedItems >= maxItems) break;
                checkedItems++;
                Console.WriteLine($"Validação: {checkedItems:N0} itens processados nesta tentativa; conferindo {row.Key[..12]}.");
                try
                {
                    if (row.FastIdentityHash is not null && row.RepairToken is null && row.State!=ArchiveItemState.DeleteStarted)
                    {
                        if (!includeFast) continue;
                        // Once full immutable evidence exists, the fresh deep
                        // proof below supersedes the provisional native index.
                        if(row.StageMsgPath is null || row.TargetSnapshotHash is null)HydrateFastEvidence(job, row, volume);
                    }
                    using var evidenceLease=new ArchiveEvidenceLease(row);
                    SourceLocation? location=null;
                    if(autoMove)
                        location=backend.LocateSource(row) ?? throw new ArchiveReviewException("Original não localizado para movimento automático; cópia preservada.");
                    var proof = Validate(job, row, volume, config, disk, autoMove && row.State!=ArchiveItemState.Copied,location);
                    // A successful read does not reset movement failures until
                    // this item's movement also completes.
                    if (!proof.Result.AllowedToRemove)
                    {
                        if (!proof.Result.SourceUnchanged || !proof.Result.TargetUnchanged || !proof.Result.SourceMatchesTarget) Review(row, string.Join(",", proof.Result.Reasons));
                        else { row.Error = string.Join(",", proof.Result.Reasons); ledger.UpdateItem(row); }
                        continue;
                    }
                    if(row.State!=ArchiveItemState.DeleteStarted)
                    { row.State = ArchiveItemState.Verified; row.Error = null; ledger.UpdateItem(row); }
                    if(autoMove)MoveValidatedOriginal(job,row,proof.Source,proof.Target,location!);
                    consecutiveAutomaticFailures=0;
                }
                catch (ArchiveInjectedCrashException) { throw; }
                catch (ArchiveReviewException ex) { Review(row, ex.Message); }
                catch (InvalidDataException ex) when (ex.Message.StartsWith("MSG sem stream de corpo reconhecido", StringComparison.Ordinal))
                { Review(row, ex.Message); }
                catch (InvalidOperationException ex) when (ex.Message.StartsWith("Anexo especial não suportado", StringComparison.Ordinal))
                { Review(row, ex.Message); }
                catch (Exception ex)
                {
                    row.Error = "verification-pending:" + ex.GetType().Name + ":" + ex.Message; ledger.UpdateItem(row);
                    if (maxItems != int.MaxValue) Console.Error.WriteLine(ex.ToString());
                    if((autoMove || workflowSelection is not null) && ++consecutiveAutomaticFailures>=3)
                    {Console.Error.WriteLine("Movimento automático interrompido após três falhas consecutivas. Confira o Outlook clássico e retome a mesma execução.");break;}
                }
                finally { workflowProgress?.Invoke(); }
            }
        }
        Status(jobId);
    }
    private void MoveValidatedOriginal(ArchiveJob job,ArchiveItem row,ArchiveSnapshot source,ArchiveSnapshot target,SourceLocation location)
    {
        if(row.State==ArchiveItemState.Verified)
        {
            SaveRemovalProof(job,row,source.Hash,target.Hash);
            row.State=ArchiveItemState.DeleteStarted;row.Error=null;ledger.UpdateItem(row);
            Fault?.Invoke("after-delete-intent",row);
        }
        if(row.State!=ArchiveItemState.DeleteStarted)throw new ArchiveReviewException("Movimento automático sem intenção válida.");
        // The fresh validation above is reused. The backend still checks the
        // change key and full relevant source information after adding its marker.
        if(!location.InDeletedItems)location=backend.MoveToDeletedItems(row,source.ChangeKey);
        Fault?.Invoke("after-trash-move-before-checkpoint",row);
        if(!location.InDeletedItems)throw new ArchiveReviewException("O Outlook não confirmou o destino Itens Excluídos.");
        row.RemovalEntryId=location.EntryId;row.RemovalFolderId=location.FolderId;row.RemovalRecordKey=location.RecordKey;
        row.State=ArchiveItemState.InDeletedItems;row.Error=null;ledger.UpdateItem(row);
        Console.WriteLine($"Validação: {row.Key[..12]} validado e movido para Itens Excluídos.");
    }
    internal void Remove(string jobId, bool removeFromMailbox, bool allowRetainedBackup = false)
    {
        if(backend is OutlookArchiveBackend lookup)lookup.BeginSourceLookupBatch();
        (backend as IArchiveBulkFastBatchBackend)?.ClearFastIdentityCache();
        var job = ledger.GetJob(jobId); var config = Envelope(job).Config;
        if (allowRetainedBackup) config.AllowRetainedMsgBackupForRemoval = true;
        if (!NativeWorkersIdle(jobId)) return;
        int consecutiveComFailures = 0;
        bool Eligible(ArchiveItem x) => x.State is ArchiveItemState.Verified or ArchiveItemState.DeleteStarted or ArchiveItemState.PurgeStarted ||
            removeFromMailbox && x.State == ArchiveItemState.InDeletedItems;
        if (!removeFromMailbox)
            Console.WriteLine($"Retomada: {ledger.GetItems(jobId).Count(x => x.State == ArchiveItemState.InDeletedItems):N0} mensagens já estão em Itens Excluídos e serão puladas.");
        foreach (var volume in ledger.GetVolumes(jobId))
        {
            if (consecutiveComFailures >= 3) break;
            if (volume.State != ArchiveVolumeState.Ready) continue;
            if (!ledger.GetItems(jobId).Any(x => x.VolumeNumber == volume.Number && Eligible(x))) continue;
            DiskRoundtrip disk;
            try
            {
                disk = backend.ReopenVolume(volume); volume.ObservedBytes = backend.MeasureVolume(volume);
                if (volume.ObservedBytes > config.HardLimitBytes - config.ReserveBytes) volume.State = ArchiveVolumeState.ReviewRequired;
                ledger.UpdateVolume(volume);
                if (volume.State != ArchiveVolumeState.Ready) continue;
            }
            catch (Exception ex)
            {
                foreach (var pending in ledger.GetItems(jobId).Where(x => x.VolumeNumber == volume.Number && Eligible(x)))
                { pending.Error = "volume-reopen-pending:" + ex.Message; ledger.UpdateItem(pending); }
                Console.Error.WriteLine("Remoção da parte bloqueada: " + ex.Message); continue;
            }
            foreach (var row in ledger.GetItems(jobId).Where(x => x.VolumeNumber == volume.Number &&
                Eligible(x)))
            {
                try
                {
                    using var evidenceLease=new ArchiveEvidenceLease(row);
                    var location = backend.LocateSource(row);
                    if (location is null)
                    {
                        if (row.State != ArchiveItemState.PurgeStarted) throw new ArchiveReviewException("Origem ausente sem intenção durável de remoção final.");
                        ValidateAbsentAfterPurge(job, row, volume, config, disk);
                        row.State = ArchiveItemState.Removed; row.Error = "reconciled-absent-after-durable-delete-intent"; ledger.UpdateItem(row); continue;
                    }
                    var proof = Validate(job, row, volume, config, disk, true, location);
                    if (!proof.Result.AllowedToRemove) throw new ArchiveReviewException(string.Join(",", proof.Result.Reasons));
                    if (row.State == ArchiveItemState.Verified)
                    {
                        row.Error = null;
                        SaveRemovalProof(job, row, proof.Source.Hash, proof.Target.Hash);
                        row.State = ArchiveItemState.DeleteStarted; ledger.UpdateItem(row); Fault?.Invoke("after-delete-intent", row);
                    }
                    bool movedNow = false;
                    if (row.State == ArchiveItemState.DeleteStarted)
                    {
                        row.Error = null;
                        if (!location.InDeletedItems)
                        { location = backend.MoveToDeletedItems(row, proof.Source.ChangeKey); movedNow = true; }
                        Fault?.Invoke("after-trash-move-before-checkpoint", row);
                        row.RemovalEntryId = location.EntryId; row.RemovalFolderId = location.FolderId; row.RemovalRecordKey = location.RecordKey;
                        row.State = ArchiveItemState.InDeletedItems; ledger.UpdateItem(row);
                    }
                    if (!removeFromMailbox) { consecutiveComFailures = 0; continue; }
                    if (movedNow) proof = Validate(job, row, volume, config, disk, true, location);
                    if (!proof.Result.AllowedToRemove || !location.InDeletedItems) throw new ArchiveReviewException("Conferência final em Itens Excluídos não passou.");
                    SaveRemovalProof(job, row, proof.Source.Hash, proof.Target.Hash);
                    row.State = ArchiveItemState.PurgeStarted; ledger.UpdateItem(row); Fault?.Invoke("after-purge-intent", row);
                    backend.DeleteFromDeletedItems(row, location, proof.Source.ChangeKey);
                    Fault?.Invoke("after-delete-before-checkpoint", row);
                    if (backend.LocateSource(row) is not null) throw new ArchiveReviewException("O item ainda existe após a remoção; não será repetida às cegas.");
                    row.State = ArchiveItemState.Removed; row.Error = null; ledger.UpdateItem(row);
                    consecutiveComFailures = 0;
                }
                catch (ArchiveInjectedCrashException) { throw; }
                catch (ArchiveReviewException ex) { Review(row, ex.Message); }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    row.Error = $"removal-interrupted:COMException:0x{ex.HResult:X8}:{ex.Message}";
                    ledger.UpdateItem(row);
                    Console.Error.WriteLine($"Remoção pendente: {row.Key[..12]} — {row.Error}. O checkpoint foi preservado; nenhuma repetição automática deste item nesta tentativa.");
                    if (++consecutiveComFailures >= 3)
                    {
                        Console.Error.WriteLine("Remoção interrompida após três falhas consecutivas do Outlook. Retome após conferir a conexão e a disponibilidade da caixa.");
                        break;
                    }
                }
                catch (Exception ex)
                {
                    row.Error = $"removal-interrupted:{ex.GetType().Name}:0x{ex.HResult:X8}:{ex.Message}";
                    ledger.UpdateItem(row);
                    Console.Error.WriteLine($"Remoção interrompida: {row.Key[..12]} — {row.Error}");
                    break;
                }
            }
        }
        Status(jobId);
    }
    private (ValidationResult Result, ArchiveSnapshot Source, ArchiveSnapshot Target) Validate(ArchiveJob job, ArchiveItem row, ArchiveVolume volume,
        ArchiveConfig config, DiskRoundtrip disk, bool forRemoval, SourceLocation? location = null)
    {
        CheckBackupFile(job, row);
        var source = backend.SourceSnapshot(row, location, Evidence(job, row, "source-current"));
        var backup = backend.BackupSnapshot(row, Evidence(job, row, "backup-check"));
        var identity = Unique(volume, row);
        var target = backend.TargetSnapshot(volume, row, identity, Evidence(job, row, "target-current"));
        return EvaluatePrepared(job,row,volume,config,disk,forRemoval,location,source,backup,target,identity:identity);
    }
    private (ValidationResult Result, ArchiveSnapshot Source, ArchiveSnapshot Target) EvaluatePrepared(ArchiveJob job,ArchiveItem row,ArchiveVolume volume,
        ArchiveConfig config,DiskRoundtrip disk,bool forRemoval,SourceLocation? location,ArchiveSnapshot source,ArchiveSnapshot backup,ArchiveSnapshot target,bool immediateTrash=false,DestinationIdentity? identity=null)
    {
        var original=Original(job,row);var baseline=TargetBaseline(job,row);
        identity??=new DestinationIdentity(row.TargetStoreId!,row.TargetFolderId!,row.TargetEntryId!);
        bool backupValid = RecordedBackupMatches(job, row, backup);
        bool expectedFolder = location is null || location.FolderId == row.SourceFolderId ||
            (location.InDeletedItems && row.State is ArchiveItemState.DeleteStarted or ArchiveItemState.InDeletedItems or ArchiveItemState.PurgeStarted);
        var request = new ValidationRequest
        {
            SourceOriginal = original, SourceCurrent = source.Fields, TargetBaseline = baseline, TargetCurrent = target.Fields,
            SourceRelevant = backup.RelevantFields, SourceCurrentRelevant = source.RelevantFields, TargetRelevant = target.RelevantFields,
            TargetBaselineRelevant = TargetRelevantBaseline(job, row),
            SourceStoreId = row.SourceStoreId, SourceEntryId = location?.EntryId ?? row.SourceEntryId,
            TargetStoreId = identity.StoreId, TargetEntryId = identity.EntryId,
            DestinationIsVerifiedLocalPst = true, DestinationBelongsToJob = true,
            DestinationReopenedAndExclusiveRead = disk.Reopened && disk.ExclusiveDiskRead,
            DestinationReopenedAndRevalidated = disk.Reopened,
            DestinationStoredItemReadBack=immediateTrash,
            TrashOnly=immediateTrash,
            IndependentMsgBackupValidated = config.AllowRetainedMsgBackupForRemoval && backupValid,
            StagingMsgHashMatches = backupValid, SourceStillInExpectedFolder = expectedFolder,
            SourceIsSupportedSentMessage = true, ImportIdentityIsUnique = true,
            DurablePhaseAllowsRemoval = !forRemoval || row.State is ArchiveItemState.Verified or ArchiveItemState.DeleteStarted or ArchiveItemState.InDeletedItems or ArchiveItemState.PurgeStarted
        };
        var result = ArchiveValidation.Evaluate(request);
        ArchiveFiles.WriteNew(Path.Combine(Path.GetDirectoryName(target.MessagePath)!, "validation.json"), new
        { Utc = DateTimeOffset.UtcNow, result, disk, BackupValidated = backupValid, BackupFallbackExplicitlyEnabled = config.AllowRetainedMsgBackupForRemoval,ImmediateTrashReadback=immediateTrash });
        return (result, source, target);
    }
    private void SaveRemovalProof(ArchiveJob job, ArchiveItem row, string sourceHash, string targetHash)
    {
        string path = Path.Combine(Evidence(job, row, "removal-proof"), "proof.json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        ArchiveFiles.WriteNew(path, new RemovalProof(row.Key, sourceHash, targetHash, row.StageMsgSha256!, DateTimeOffset.UtcNow, row.SnapshotHash, row.TargetSnapshotHash));
        row.RemovalProofPath = path; row.RemovalProofSha256 = ArchiveFiles.Sha(path);
    }
    private void ValidateAbsentAfterPurge(ArchiveJob job, ArchiveItem row, ArchiveVolume volume, ArchiveConfig config, DiskRoundtrip disk)
    {
        CheckBackupFile(job, row);
        string path = ArchiveFiles.Child(job.RootDirectory, row.RemovalProofPath ?? throw new ArchiveReviewException("Prova anterior ausente."));
        if (ArchiveFiles.Sha(path) != row.RemovalProofSha256) throw new ArchiveReviewException("Prova anterior alterada.");
        var proof = JsonSerializer.Deserialize<RemovalProof>(File.ReadAllText(path))!;
        var backup = backend.BackupSnapshot(row, Evidence(job, row, "backup-after-delete"));
        var target = backend.TargetSnapshot(volume, row, Unique(volume, row), Evidence(job, row, "target-after-delete"));
        var baseline = TargetBaseline(job, row); var baselineRelevant = TargetRelevantBaseline(job, row);
        bool targetIntact = target.Hash == row.TargetSnapshotHash ||
            EmbeddedMessageEvidence.SameRepresentations(baseline, target.Fields, baselineRelevant, target.RelevantFields) &&
            RelevantMessageValidation.Equivalent(new(baseline, "unused", "unused", null, baselineRelevant), target);
        if (proof.Key != row.Key || (proof.SourceBaselineHash ?? proof.SourceHash) != row.SnapshotHash || (proof.TargetBaselineHash ?? proof.TargetHash) != row.TargetSnapshotHash ||
            proof.StageSha != row.StageMsgSha256 || !RecordedBackupMatches(job, row, backup) || !targetIntact ||
            !RelevantMessageValidation.Equivalent(backup, target) ||
            !disk.Reopened || (!disk.ExclusiveDiskRead && !config.AllowRetainedMsgBackupForRemoval))
            throw new ArchiveReviewException("Reconciliação da remoção exige a cópia e o backup íntegros.");
    }
    private Dictionary<string, string> Original(ArchiveJob job, ArchiveItem row)
    {
        string path = ArchiveFiles.Child(job.RootDirectory, Path.Combine(Path.GetDirectoryName(row.StageMsgPath!)!, "field-hashes.json"));
        var fields = ArchiveFiles.ReadFields(path);
        if (ArchiveFiles.HashFields(fields) != row.SnapshotHash) throw new ArchiveReviewException("Snapshot original foi alterado."); return fields;
    }
    private Dictionary<string, string> TargetBaseline(ArchiveJob job, ArchiveItem row)
    {
        string path = ArchiveFiles.Child(job.RootDirectory, row.TargetFieldsPath ?? throw new ArchiveReviewException("Baseline do destino ausente."));
        var fields = ArchiveFiles.ReadFields(path);
        if (ArchiveFiles.HashFields(fields) != row.TargetSnapshotHash) throw new ArchiveReviewException("Baseline do destino foi alterado."); return fields;
    }
    private void CheckBackupFile(ArchiveJob job, ArchiveItem row)
    {
        string path = ArchiveFiles.Child(job.RootDirectory, row.StageMsgPath ?? throw new ArchiveReviewException("Backup MSG não registrado."));
        if (Path.GetExtension(path) != ".msg" || ArchiveFiles.Sha(path) != row.StageMsgSha256) throw new ArchiveReviewException("Backup MSG ausente ou alterado.");
    }
    private DestinationIdentity Unique(ArchiveVolume volume, ArchiveItem row)
    {
        var found = backend.FindToken(volume, row);
        if (found.Matches.Count != 1) throw new ArchiveReviewException("Token de importação ausente ou duplicado no destino.");
        return found.Matches[0];
    }
    private ArchiveVolume Volume(ArchiveItem row) => ledger.GetVolumes(row.JobId).Single(x => x.Number == row.VolumeNumber);
    private string Evidence(ArchiveJob job, ArchiveItem row, string stage)
    {
        // Keep COM file paths below Outlook's legacy path-length boundary.
        // The full identity stays in SQLite; paths are never used as identity.
        string path = Path.Combine(job.RootDirectory, "i", row.Key[..16], stage + "-" + Guid.NewGuid().ToString("N")[..12]);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!); return path;
    }
    private void Review(ArchiveItem row, string error)
    { row.State = ArchiveItemState.ReviewRequired; row.Error = error; ledger.UpdateItem(row); Console.Error.WriteLine("Revisão pendente: " + row.Key[..12] + " " + error); }
    private void CheckSpace(long estimate, ArchiveConfig config)
    {
        long needed = checked(Math.Max(estimate, 0) * 4 + config.ReserveBytes);
        if (new DriveInfo(Path.GetPathRoot(root)!).AvailableFreeSpace < needed) throw new IOException("Espaço insuficiente para exportação, cópia e margem.");
    }
    internal object Status(string jobId)
    {
        if(workflowSelection is not null)return new{JobId=jobId};
        var job = ledger.GetJob(jobId); var envelope = Envelope(job); var items = ledger.GetItems(jobId);
        var plan = JsonSerializer.Deserialize<List<SourceDescriptor>>(File.ReadAllText(Path.Combine(job.RootDirectory, "source-plan.json")))!;
        int previous = plan.Count(x => ledger.GetItem(ArchiveLedger.ComputeItemKey(x.StoreId, x.EntryId))?.JobId != jobId);
        var result = new { JobId = jobId, RootDirectory = root, States = items.GroupBy(x => x.State.ToString()).ToDictionary(x => x.Key, x => x.Count()),
            Volumes = ledger.GetVolumes(jobId).Select(x => new { x.Number, x.PstPath, x.ObservedBytes, State = x.State.ToString() }).ToArray(),
            PreviouslyRegisteredInOtherJobs = previous, RetainedMsgBackupForRemoval = envelope.Config.AllowRetainedMsgBackupForRemoval, envelope.Config.PreserveFolderStructure,
            CutoffDateInclusive = envelope.Config.CutoffDateInclusive, CutoffTimeZoneId = envelope.Config.CutoffTimeZoneId,
            Scope = "Itens do plano congelado; não certifica completude no servidor ou quota liberada.", InterfaceImplemented = true };
        string path = Path.Combine(job.RootDirectory, "status-" + Guid.NewGuid().ToString("N") + ".json"); ArchiveFiles.WriteNew(path, result);
        Console.WriteLine(JsonSerializer.Serialize(result, ArchiveFiles.JsonOptions)); return result;
    }
    internal List<ArchiveItem> Items(string job) => ledger.GetItems(job);
    internal bool AutomaticMovementEnabled(string job)=>ledger.WorkflowManifest(job)?.Move ?? Envelope(ledger.GetJob(job)).Config.MoveValidatedToDeletedItems;
    internal bool HasWorkflow(string job)=>ledger.WorkflowManifest(job) is not null;
    internal List<ArchiveVolume> Volumes(string job) => ledger.GetVolumes(job);
    public void Dispose() { ledger.Dispose(); lease.Dispose(); }
    private sealed record RemovalProof(string Key, string SourceHash, string TargetHash, string StageSha, DateTimeOffset Utc, string? SourceBaselineHash = null, string? TargetBaselineHash = null);
}
