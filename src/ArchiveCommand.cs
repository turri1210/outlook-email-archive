using System.Text.Json;

internal static class ArchiveCommand
{
    internal static int Execute(string[] args)
    {
        if (args.Length == 0) return 2;
        try
        {
            bool move=args.Contains("--move-validated-to-trash",StringComparer.Ordinal);
            bool reconcilePending=args.Contains("--reconcile-pending",StringComparer.Ordinal);
            if(reconcilePending)
            {
                if(args[0]!="archive-workflow-run")throw new ArgumentException("Reconciliação por lotes é válida apenas no fluxo de arquivamento.");
                args=args.Where(x=>x!="--reconcile-pending").ToArray();
            }
            bool keep=args.Contains("--keep-originals",StringComparer.Ordinal);
            if(move||keep)
            {
                if(move&&keep || args[0] is not ("archive-run" or "archive-verify" or "archive-workflow-run" or "archive-workflow-prepare"))throw new ArgumentException("Opção de movimento automático inválida.");
                args=args.Where(x=>x is not ("--move-validated-to-trash" or "--keep-originals")).ToArray();
            }
            bool? autoMove=move ? true : keep ? false : null;
            if(args[0]=="archive-workflow-self-test" && args.Length==1)return ArchiveWorkflowTests.Run();
            if(args[0]=="archive-storage-self-test" && args.Length==1)return ArchiveStorageTests.Run();
            if(args[0]=="archive-source-index-audit" && args.Length==3)
            {
                using var auditLock=new FileStream(Path.Combine(args[1],"archive.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
                using var auditLedger=new ArchiveLedger(Path.Combine(args[1],"archive.sqlite"),true);
                using var auditBackend=new OutlookArchiveBackend();auditBackend.BeginSourceLookupBatch();
                var auditRows=auditLedger.GetItems(args[2]).Where(x=>x.State==ArchiveItemState.ReviewRequired).Take(100).ToList();
                int found=0,missing=0,blocked=0;var watch=System.Diagnostics.Stopwatch.StartNew();
                foreach(var row in auditRows)
                {
                    try{if(auditBackend.LocateSource(row) is null)missing++;else found++;}
                    catch(ArchiveReviewException){blocked++;}
                }
                Console.WriteLine(JsonSerializer.Serialize(new{Checked=auditRows.Count,Found=found,Missing=missing,Blocked=blocked,ElapsedSeconds=watch.Elapsed.TotalSeconds,ReadOnly=true}));return 0;
            }
            if(args[0]=="archive-maintain-storage" && args.Length==3)
            {
                ArchiveFiles.ValidateLocalDirectory(args[1]);
                using var storageLock=new FileStream(Path.Combine(args[1],"archive.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
                using var storageLedger=new ArchiveLedger(Path.Combine(args[1],"archive.sqlite"));var storageJob=storageLedger.GetJob(args[2]);
                if(Path.GetFullPath(storageJob.RootDirectory)!=Path.GetFullPath(Path.Combine(args[1],"jobs",storageJob.Id)))throw new ArchiveReviewException("Execução fora do histórico selecionado.");
                bool goal=storageLedger.WorkflowManifest(args[2])?.Move??JsonSerializer.Deserialize<ArchivePlanEnvelope>(storageJob.ConfigJson)!.Config.MoveValidatedToDeletedItems;
                Console.WriteLine(JsonSerializer.Serialize(ArchiveStorageMaintenance.Run(storageLedger,storageJob,storageLedger.GetItems(args[2]),goal,true),ArchiveFiles.JsonOptions));return 0;
            }
            bool pendingOnly = args.Contains("--pending-only", StringComparer.Ordinal);
            if (pendingOnly)
            {
                if (args[0] != "archive-verify") throw new ArgumentException("--pending-only é válido apenas para archive-verify.");
                args = args.Where(x => x != "--pending-only").ToArray();
            }
            bool allowRetainedBackup = args.LastOrDefault() == "--allow-retained-msg-backup";
            if (allowRetainedBackup)
            {
                if (args[0] is not ("archive-run" or "archive-verify" or "archive-workflow-run" or "archive-remove" or "archive-resolve-pending")) throw new ArgumentException("A autorização de backup MSG é válida apenas para cópia, validação e remoção.");
                args = args[..^1];
                Console.WriteLine("Autorização explícita nesta operação: permitir backup MSG íntegro retido como prova adicional. O plano original permanece preservado.");
            }
            if(args[0]=="archive-workflow-prepare" && args.Length==3)
            {
                using var workflowLock=new FileStream(Path.Combine(args[1],"archive.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
                using var workflowLedger=new ArchiveLedger(Path.Combine(args[1],"archive.sqlite"));
                WorkflowPlan.Ensure(workflowLedger,workflowLedger.GetJob(args[2]),autoMove??true,recoverTransient:true);
                Console.WriteLine(JsonSerializer.Serialize(workflowLedger.WorkflowStatus(args[2]),ArchiveFiles.JsonOptions));return 0;
            }
            if(args[0]=="archive-workflow-run" && args.Length==3)
            {
                using var workflowBackend=new OutlookArchiveBackend();using var workflowEngine=new ArchiveEngine(args[1],workflowBackend);
                return workflowEngine.ContinueWorkflow(args[2],autoMove??workflowEngine.AutomaticMovementEnabled(args[2]),allowRetainedBackup,reconcilePending:reconcilePending);
            }
            if (args[0] == "archive-ndr-audit-pilot" && args.Length == 2)
            { DeliveryReportPilotAudit.Run(args[1]); return 0; }
            if (args[0] == "archive-enable-delivery-reports" && args.Length == 3)
            { using var backend=new OutlookArchiveBackend(); using var engine=new ArchiveEngine(args[1],backend); engine.PrepareDeliveryReports(args[2]); return 0; }
            if (args[0] == "archive-ndr-pilot" && args.Length == 4)
            { using var backend=new OutlookArchiveBackend(); return backend.DeliveryReportPilot(args[1],args[2],args[3]); }
            if (args[0] == "archive-fast-recovery-inspect" && args.Length == 4)
            {
                using var backend = new OutlookArchiveBackend();
                backend.InspectFastRecovery(args[1], args[2], args[3]); return 0;
            }
            if (args[0] == "archive-recovery-inspect" && args.Length == 4)
            {
                using var backend = new OutlookArchiveBackend();
                backend.InspectRecoveryEvidence(args[1], args[2], args[3]); return 0;
            }
            if (args[0] == "archive-recovery-pilot-cleanup" && args.Length == 4)
            {
                using var backend = new OutlookArchiveBackend();
                backend.CleanupRecoveryPilot(args[1], args[2], args[3]); return 0;
            }
            if (args[0] == "archive-recovery-pilot" && args.Length == 4)
            {
                using var backend = new OutlookArchiveBackend();
                return backend.RecoveryPilot(args[1], args[2], args[3]);
            }
            if (args[0] == "archive-ui-sources" && args.Length == 1)
            {
                using var backend = new OutlookArchiveBackend();
                Console.WriteLine(JsonSerializer.Serialize(backend.ListSources(), ArchiveFiles.JsonOptions)); return 0;
            }
            if (args[0] is "archive-destination-check" or "archive-destination-reopen-check" && args.Length == 3)
            {
                using var read = new ArchiveLedger(Path.Combine(args[1], "archive.sqlite"), true);
                var job = read.GetJob(args[2]);
                using var backend = new OutlookArchiveBackend();
                foreach (var volume in read.GetVolumes(job.Id))
                {
                    if (!File.Exists(volume.PstPath)) throw new ArchiveReviewException("PST registrado ausente; verificação interrompida.");
                    backend.AttachVolume(volume);
                    Console.WriteLine($"Destino registrado confirmado: parte {volume.Number}, {volume.PstPath}");
                    if (args[0] == "archive-destination-reopen-check")
                    {
                        try { Console.WriteLine(JsonSerializer.Serialize(backend.ReopenVolume(volume))); }
                        catch (Exception ex) { Console.Error.WriteLine(ex.ToString()); return 3; }
                    }
                }
                return 0;
            }
            if (args[0] == "archive-store-lookup-self-test" && args.Length == 1) return RecordedStoreLookup.SelfTest();
            if (args[0] == "archive-folder-structure-smoke" && args.Length == 3)
            {
                using var backend = new OutlookArchiveBackend();
                Console.WriteLine(JsonSerializer.Serialize(backend.FolderStructureSmoke(args[1], args[2]), ArchiveFiles.JsonOptions)); return 0;
            }
            if (args[0] == "archive-removal-preparation-check" && args.Length == 3)
            {
                using var read = new ArchiveLedger(Path.Combine(args[1], "archive.sqlite"), true);
                int passed = 0;
                foreach (var row in read.GetItems(args[2]))
                {
                    string directory = Path.GetDirectoryName(row.StageMsgPath!)!;
                    string? latest = Directory.EnumerateDirectories(directory, "mark-*").OrderByDescending(x => Directory.GetLastWriteTimeUtc(x)).FirstOrDefault();
                    if (latest is null) throw new ArchiveReviewException("Evidência da preparação ausente.");
                    var fields = ArchiveFiles.ReadFields(Path.Combine(latest, "field-hashes.json"));
                    string relevantPath = Path.Combine(latest, "relevant-fields.json");
                    var relevant = File.Exists(relevantPath) ? ArchiveFiles.ReadFields(relevantPath) : null;
                    if (!ArchivedSourceEvidence.Matches(row, new(fields, Path.Combine(latest, "message.msg"), "unused", null, relevant)))
                        throw new ArchiveReviewException("A evidência de preparação ainda diverge da origem arquivada.");
                    passed++;
                }
                Console.WriteLine($"Preparação da remoção: {passed} evidências reais aprovadas; nenhum original alterado."); return 0;
            }
            if (args[0] == "archive-token-check" && args.Length == 3)
            {
                using var read = new ArchiveLedger(Path.Combine(args[1], "archive.sqlite"), true);
                using var backend = new OutlookArchiveBackend();
                var volume = read.GetVolumes(args[2]).First(); backend.AttachVolume(volume);
                var sample = read.GetItems(args[2]).First(x => x.VolumeNumber == volume.Number && x.FastIdentityHash is not null);
                var token = backend.FindToken(volume, sample);
                Console.WriteLine($"Cópias correspondentes ao identificador da amostra: {token.Matches.Count}"); return 0;
            }
            if (args[0] == "archive-content-check" && args.Length == 4)
            {
                using var read = new ArchiveLedger(Path.Combine(args[1], "archive.sqlite"), true);
                using var backend = new OutlookArchiveBackend();
                var volume = read.GetVolumes(args[2]).First(); backend.AttachVolume(volume);
                var sample = read.GetItems(args[2]).First(x => x.VolumeNumber == volume.Number && x.FastIdentityHash is not null && x.StageMsgPath is not null &&
                    File.Exists(Path.Combine(Path.GetDirectoryName(x.StageMsgPath)!, "relevant-fields.json")) &&
                    File.ReadAllText(Path.Combine(Path.GetDirectoryName(x.StageMsgPath)!, "relevant-fields.json")).Contains("html-relevant-v1", StringComparison.Ordinal));
                string evidence = Path.GetFullPath(args[3]);
                if (Directory.Exists(evidence)) throw new IOException("Diretório de diagnóstico já existe.");
                var match = backend.FindToken(volume, sample);
                if (match.Matches.Count != 1) throw new ArchiveReviewException("Cópia da amostra ambígua ou ausente.");
                var source = backend.SourceSnapshot(sample, null, Path.Combine(evidence, "source"));
                var backup = backend.BackupSnapshot(sample, Path.Combine(evidence, "backup"));
                var target = backend.TargetSnapshot(volume, sample, match.Matches[0], Path.Combine(evidence, "target"));
                bool valid = backup.MessageSha256 == sample.StageMsgSha256 && RelevantMessageValidation.Equivalent(source, target) && RelevantMessageValidation.Equivalent(backup, target);
                Console.WriteLine($"Amostra real: original, backup MSG íntegro e cópia PST correspondem: {valid}. Catálogo e originais preservados.");
                return valid ? 0 : 3;
            }
            if (args[0] == "archive-verify-sample" && args.Length == 3)
            {
                using var backend = new OutlookArchiveBackend(); using var engine = new ArchiveEngine(args[1], backend);
                engine.Verify(args[2], maxItems: 3); return 0;
            }
            if (args[0] == "archive-embedded-self-test" && args.Length == 1)
            { EmbeddedMessageEvidence.SelfTest(); return 0; }
            if (args[0] == "archive-self-test" && args.Length == 1) return ArchiveEngineTests.Run();
            if (args[0] == "archive-fast-self-test" && args.Length == 1) return ArchiveFastBatchTests.Run();
            if (args[0] == "archive-fast-smoke" && args.Length == 4) return NativeBatchSmoke.Execute(args.Skip(1).ToArray(), true);
            if (args[0] == "archive-batch-self-test" && args.Length == 1) return ArchiveBatchTests.Run();
            if (args[0] == "archive-relevant-self-test" && args.Length == 1) { RelevantMessageValidation.SelfTest(); Console.WriteLine("Validação das informações relevantes aprovada."); return 0; }
            if (args[0] == "archive-native-smoke" && args.Length == 4) return NativeBatchSmoke.Execute(args.Skip(1).ToArray());
            if (args[0] == "archive-perf-test" && args.Length == 1) return ArchivePerformanceBenchmark.Run();
            if (args[0] is "archive-check-new" or "archive-plan-new" && args.Length == 3)
            {
                using var backend = new OutlookArchiveBackend(); using var engine = new ArchiveEngine(args[1], backend);
                if (args[0] == "archive-plan-new") { Console.WriteLine("EXECUCAO=" + engine.PlanNew(args[2])); return 0; }
                var writer = Console.Out; ArrivalCheckResult result;
                try { Console.SetOut(TextWriter.Null); result = engine.CheckNew(args[2]); }
                finally { Console.SetOut(writer); }
                Console.WriteLine(JsonSerializer.Serialize(result, ArchiveFiles.JsonOptions)); return result.Error is null ? 0 : 3;
            }
            if (args[0] == "archive-plan" && args.Length == 2)
            {
                var config = JsonSerializer.Deserialize<ArchiveConfig>(File.ReadAllText(args[1])) ?? throw new ArgumentException("Configuração inválida.");
                config.Validate(); using var backend = new OutlookArchiveBackend(); using var engine = new ArchiveEngine(config.RootDirectory, backend);
                string id = engine.Plan(config); Console.WriteLine("EXECUCAO=" + id); return 0;
            }
            if (args[0] == "archive-status" && args.Length is 2 or 3)
            {
                ArchiveFiles.ValidateLocalDirectory(args[1]);
                using var ledger = new ArchiveLedger(Path.Combine(args[1], "archive.sqlite"), true);
                var jobs = ledger.GetJobs().Where(x => args.Length == 2 || x.Id == args[2]);
                Console.WriteLine(JsonSerializer.Serialize(jobs.Select(job =>
                {
                    var config = JsonSerializer.Deserialize<ArchivePlanEnvelope>(job.ConfigJson)!.Config;
                    var items = ledger.GetItems(job.Id);
                    return new { job.Id, job.CreatedUtc,
                        SourceStoreId = config.SourceStoreId,
                        SourceFolderId = config.SourceFolderId,
                        SourceLabel = config.SourceLabel,
                        ExistingPstPath = config.ExistingPstPath,
                        DestinationPstPath = config.DestinationPstPath,
                        config.MoveValidatedToDeletedItems,
                        Workflow=ledger.WorkflowStatus(job.Id),
                        CutoffDateInclusive = config.CutoffDateInclusive,
                        CutoffTimeZoneId = config.CutoffTimeZoneId,
                        RetainedMsgBackupForRemoval = config.AllowRetainedMsgBackupForRemoval, config.PreserveFolderStructure,
                        States = items.GroupBy(x => x.State.ToString()).ToDictionary(x => x.Key, x => x.Count()),
                        Errors = items.Where(x => x.Error is not null && x.State != ArchiveItemState.Removed && x.State != ArchiveItemState.BackupVerifiedSourceUnmanaged && !(x.State == ArchiveItemState.Copied && x.Error.StartsWith("Cópia rápida concluída;", StringComparison.Ordinal))).Select(x => new { Key = x.Key, x.Error }).ToArray(),
                        Volumes = ledger.GetVolumes(job.Id).Select(x => new { x.Number, x.PstPath, x.ObservedBytes, State = x.State.ToString() }).ToArray() };
                }), ArchiveFiles.JsonOptions)); return 0;
            }
            if(args[0]=="archive-inspect-remaining-types" && args.Length==4)
            {using var backend=new OutlookArchiveBackend();backend.InspectRemainingTypes(args[1],args[2],args[3]);return 0;}
            if(args[0]=="archive-restore-repair-html" && args.Length==4)
            {
                var keys=JsonSerializer.Deserialize<HashSet<string>>(File.ReadAllText(args[3])) ?? throw new ArgumentException("Seleção inválida.");
                using var backend=new OutlookArchiveBackend();using var engine=new ArchiveEngine(args[1],backend);engine.RestoreRepairHtml(args[2],keys);return 0;
            }
            if (args[0] == "archive-resolve-pending" && args.Length == 4)
            {
                var keys = JsonSerializer.Deserialize<HashSet<string>>(File.ReadAllText(args[3])) ?? throw new ArgumentException("Lista inválida.");
                using var backend = new OutlookArchiveBackend { RetainComparisonSnapshots = true }; using var engine = new ArchiveEngine(args[1],backend);
                if (keys.Count == 0 || keys.Any(key => !engine.Items(args[2]).Any(row => row.Key == key))) throw new ArgumentException("Item fora da execução selecionada.");
                engine.ResolvePending(args[2],keys,allowRetainedBackup);
                return engine.Items(args[2]).Any(row => keys.Contains(row.Key) && row.State == ArchiveItemState.ReviewRequired) ? 3 : 0;
            }
            if (args[0] == "archive-reconcile-selected" && args.Length == 4)
            {
                var keys = JsonSerializer.Deserialize<HashSet<string>>(File.ReadAllText(args[3])) ?? throw new ArgumentException("Lista de itens inválida.");
                using var backend = new OutlookArchiveBackend(); using var engine = new ArchiveEngine(args[1], backend);
                if (keys.Count == 0 || keys.Any(key => !engine.Items(args[2]).Any(row => row.Key == key))) throw new ArgumentException("Item fora da execução selecionada.");
                engine.Reconcile(args[2], keys); return engine.Items(args[2]).Any(row => keys.Contains(row.Key) && row.State == ArchiveItemState.ReviewRequired) ? 3 : 0;
            }
            bool mutation = args[0] is "archive-run" or "archive-verify" or "archive-remove" or "archive-reconcile";
            if (mutation && args.Length is 3 or 4 or 5)
            {
                bool remove = args[0] == "archive-remove";
                if (remove && (args.Length < 4 || args[3] is not ("deleted-items" or "mailbox")))
                    throw new ArgumentException("Remoção exige modo explícito: deleted-items ou mailbox.");
                if (!remove && args.Length > 4) throw new ArgumentException("Argumentos inválidos.");
                using var backend = new OutlookArchiveBackend(); using var engine = new ArchiveEngine(args[1], backend);
                string? test = args.LastOrDefault(x => x.StartsWith("--test-exit=", StringComparison.Ordinal));
                if (test is not null)
                {
                    using var read = new ArchiveLedger(Path.Combine(args[1], "archive.sqlite"), true);
                    var envelope = JsonSerializer.Deserialize<ArchivePlanEnvelope>(read.GetJob(args[2]).ConfigJson)!;
                    if (envelope.Config.SyntheticSourcePstPath is null) throw new ArgumentException("Interrupção artificial só é permitida com origem PST artificial declarada.");
                    backend.Discover(envelope.Config); // Confirms the configured store is that local PST.
                    string point = test["--test-exit=".Length..];
                    if (point is not ("after-import-before-checkpoint" or "after-trash-move-before-checkpoint" or "after-delete-before-checkpoint" or "after-import-intent" or
                        "after-batch-copy-intent" or "after-batch-copy-before-checkpoint" or "after-batch-matched" or "after-batch-mark-before-checkpoint" or "after-batch-finalize-before-checkpoint"))
                        throw new ArgumentException("Ponto de teste inválido.");
                    engine.Fault = (phase, _) => { if (phase == point) Environment.Exit(86); };
                }
                else if (!remove && args.Length != 3 || remove && args.Length != 4)
                    throw new ArgumentException("Argumentos adicionais não reconhecidos.");
                if (args[0] == "archive-run")
                {
                    if(engine.HasWorkflow(args[2]))return engine.ContinueWorkflow(args[2],autoMove??engine.AutomaticMovementEnabled(args[2]),allowRetainedBackup);
                    engine.Copy(args[2]);
                    if(autoMove ?? engine.AutomaticMovementEnabled(args[2]))engine.Verify(args[2],allowRetainedBackup:allowRetainedBackup,pendingOnly:true,moveValidatedToDeletedItems:true);
                }
                else if (args[0] == "archive-verify") engine.Verify(args[2], allowRetainedBackup: allowRetainedBackup, pendingOnly: pendingOnly,moveValidatedToDeletedItems:autoMove);
                else if (args[0] == "archive-reconcile") engine.Reconcile(args[2]);
                else engine.Remove(args[2], args[3] == "mailbox", allowRetainedBackup);
                var rows = engine.Items(args[2]);
                return rows.Any(x => (x.Error is not null && x.State != ArchiveItemState.Removed) || x.State is ArchiveItemState.ReviewRequired or ArchiveItemState.Planned or ArchiveItemState.Exported or ArchiveItemState.ImportStarted or ArchiveItemState.Copied or ArchiveItemState.DeleteStarted or ArchiveItemState.PurgeStarted or ArchiveItemState.SkippedDraft or ArchiveItemState.SkippedUnsupported || (remove && args[3] == "mailbox" && x.State == ArchiveItemState.InDeletedItems)) ? 3 : 0;
            }
            Console.Error.WriteLine("Comandos: archive-plan <config.json>; archive-run|archive-verify <raiz> <execucao>; archive-remove <raiz> <execucao> deleted-items|mailbox; archive-status <raiz> [execucao]; archive-self-test.");
            return 2;
        }
        catch (Exception ex) { Console.Error.WriteLine("Arquivamento interrompido: " + ex.GetType().Name + ": " + ex.Message); if (args[0] == "archive-destination-reopen-check") Console.Error.WriteLine(ex.ToString()); return 1; }
    }
}
