using System.Text.Json;

// A separate synthetic provider deliberately counts every expensive snapshot.
// Nothing in this suite connects to Outlook or the user's catalog.
internal static class ArchiveFastBatchTests
{
    internal static int Run()
    {
        string root = Path.Combine(Path.GetTempPath(), "arc-fast-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var tests = new (string Name, Action<string> Test)[]
        {
            ("copy-without-msg", CopyWithoutMsg),
            ("automatic-trash-resume",AutomaticTrashResume),
            ("workflow-native-lots",WorkflowNativeLots),
            ("workflow-native-revalidation",WorkflowNativeRevalidation),
            ("ndr-full-native-evidence", DeliveryReportsUseFullNativeEvidence),
            ("reconcile-without-deferred-msg", ReconcileWithoutDeferredMsg),
            ("reconcile-native-date-precision", ReconcileNativeDatePrecision),
            ("completed-batch-deep-review-reconcile", CompletedBatchReconcile),
            ("bulk-finalize-checkpoint-resume", BulkFinalizeResume),
            ("bulk-finalize-partial-resume", BulkFinalizePartial),
            ("frozen-arrivals", FrozenArrivals),
            ("resume-after-copy", p => Resume(p, "after-batch-copy-before-checkpoint")),
            ("resume-after-match", p => Resume(p, "after-batch-matched")),
            ("resume-after-token", p => Resume(p, "after-batch-mark-before-checkpoint")),
            ("resume-after-finalize", p => Resume(p, "after-batch-finalize-before-checkpoint")),
            ("partial-does-not-repeat", PartialDoesNotRepeat),
            ("pending-batch-does-not-block-rest", PendingDoesNotBlockRest),
            ("metadata-change-pending", MetadataChangePending),
            ("partial-reconcile-preserves-adopted", PartialReconcilePreservesAdopted),
            ("fresh-verification-after-copy", FreshVerificationAfterCopy),
            ("bulk-cache-scope-and-cleanup", BulkCacheScopeAndCleanup),
            ("deep-body-difference-blocks-removal", DeepBodyDifferenceBlocksRemoval)
        };
        try
        {
            foreach (var (name, test) in tests)
            {
                string scenario = Path.Combine(root, name); Directory.CreateDirectory(scenario); test(scenario);
                Console.WriteLine("Fluxo rápido: " + name + " aprovado.");
            }
            Console.WriteLine($"Fluxo rápido: {tests.Length} cenários independentes aprovados."); return 0;
        }
        finally
        {
            string full = Path.GetFullPath(root);
            string temp = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (full.StartsWith(temp, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(full).StartsWith("arc-fast-", StringComparison.Ordinal))
                Directory.Delete(full, true);
        }
    }
    private static ArchiveConfig Config(string root) => new()
    {
        RootDirectory = root, SourceStoreId = "source", SourceFolderId = "folder", IncludeSubfolders = true,
        TargetVolumeBytes = 128 * ArchiveVolumePolicy.MiB, HardLimitBytes = 256 * ArchiveVolumePolicy.MiB, ReserveBytes = ArchiveVolumePolicy.MiB
    };
    private static FastFake Provider(string root) => new(Path.Combine(root, "provider.json"));
    private static void Check(bool condition, string message) { if (!condition) throw new Exception("Fluxo rápido: " + message); }
    private static void Intact(FastFake backend) => Check(backend.Inner.Sources.All(x => !x.Removed && x.Folder != "trash") && backend.Inner.DeleteCalls == 0, "Originais foram alterados/removidos sem autorização de validação.");
    private static void DeliveryReportsUseFullNativeEvidence(string root)
    {
        using var backend=Provider(root);
        foreach(var source in backend.Inner.Sources)
        {
            source.Fields["MessageClass"]=FakeArchiveBackend.Hash(DeliveryReportEvidence.MessageClass);
            foreach(string key in DeliveryReportEvidence.RequiredFields)source.Fields[key]=FakeArchiveBackend.Hash(key=="Report-Version"?"ndr-v1":key);
        }
        backend.Inner.Save();using var engine=new ArchiveEngine(root,backend);string job=engine.Plan(Config(root));engine.Copy(job);
        Check(backend.NativeCalls==1 && backend.ExpensiveCalls>0 && engine.Items(job).All(x=>x.StageMsgPath is not null && x.FastIdentityHash is null && x.State==ArchiveItemState.Copied),"NDR entrou no caminho rápido sem prova integral.");
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))Check(ledger.GetBatches(job).All(x=>!x.FastCopy),"Lote NDR foi rotulado como rápido.");
        Check(!backend.RefreshCalls.SelectMany(x=>x).Any(),"NDR foi enviado à identificação rápida de e-mails comuns.");
        engine.Copy(job);engine.Verify(job);
        Check(backend.NativeCalls==1 && engine.Items(job).All(x=>x.State==ArchiveItemState.Verified),"Retomada NDR duplicou cópia ou não validou.");Intact(backend);
    }
    private static void WorkflowNativeRevalidation(string root)
    {
        using var backend=Provider(root);backend.Inner.EmitRelevantMetadata=true;using var engine=new ArchiveEngine(root,backend);
        string job=engine.Plan(Config(root));engine.Copy(job);engine.Verify(job);
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
        {
            int index=0;foreach(var row in ledger.GetItems(job))
            {row.State=ArchiveItemState.ReviewRequired;row.Error=index++<4?"Identificação nativa da mensagem pendente: message unavailable":"source_changed_or_evidence_invalid";ledger.UpdateItem(row);}
            WorkflowPlan.Ensure(ledger,ledger.GetJob(job),true,recoverTransient:true);
            Check(ledger.GetItems(job).Count(x=>x.State==ArchiveItemState.Copied)==4,"Falha transitória não foi reenfileirada ou diferença real foi aprovada.");
        }
        backend.UnavailableNative=true;backend.Inner.Sources[0].Fields["Subject"]=FakeArchiveBackend.Hash("real change");backend.Inner.Save();
        Check(engine.ContinueWorkflow(job,true)==3 && engine.Items(job).Count(x=>x.State==ArchiveItemState.InDeletedItems)==3,"Validação completa reutilizou índice indisponível ou ignorou mudança real.");
        Check(backend.Inner.Sources[0].Folder=="folder" && backend.NativeCalls==1,"Revalidação removeu divergência ou repetiu cópia.");
    }
    private static void WorkflowNativeLots(string root)
    {
        using var backend=Provider(root);var original=backend.Inner.Sources[0];backend.Inner.Sources.Clear();
        for(int i=0;i<205;i++)backend.Inner.Sources.Add(new(){Id="native-"+i,RecordKey="record-"+i,Size=16384,Fields=new(original.Fields)});
        backend.Inner.Save();using var engine=new ArchiveEngine(root,backend);var config=Config(root);config.TargetVolumeBytes=256*ArchiveVolumePolicy.MiB;config.HardLimitBytes=512*ArchiveVolumePolicy.MiB;string job=engine.Plan(config);
        int check=0;Check(engine.ContinueWorkflow(job,true,pauseRequested:()=>++check>=2)==4,"Pausa do lote nativo não ocorreu.");
        Check(backend.NativeCalls==1 && backend.Inner.Copies.Count==100 && backend.Inner.Sources.Count(x=>x.Folder=="trash")==100,"Cópia nativa ultrapassou o escopo do lote atual.");
        Check(engine.ContinueWorkflow(job,true)==0 && backend.NativeCalls==3 && backend.Inner.Copies.Count==205,"Retomada nativa repetiu lotes ou perdeu itens.");
        using var catalog=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);
        var progress=catalog.WorkflowStatus(job)!;Check(progress.CompletedBatches==3 && progress.CompleteItems==205,"Conclusão dos lotes nativos incorreta.");
    }
    private static void AutomaticTrashResume(string root)
    {
        using var backend=Provider(root);using var engine=new ArchiveEngine(root,backend);
        string job=engine.Plan(Config(root));engine.Copy(job);
        engine.Fault=(phase,_)=>{if(phase=="after-trash-move-before-checkpoint")throw new ArchiveInjectedCrashException(phase);};
        try{engine.Verify(job,pendingOnly:true,moveValidatedToDeletedItems:true);throw new Exception("Interrupção automática não ocorreu.");}catch(ArchiveInjectedCrashException){}
        engine.Fault=null;engine.Verify(job,pendingOnly:true,moveValidatedToDeletedItems:true);
        Check(backend.NativeCalls==1 && engine.Items(job).All(x=>x.State==ArchiveItemState.InDeletedItems),"Retomada do movimento rápido repetiu cópia ou perdeu originais movidos.");
        Check(backend.Inner.DeleteCalls==0 && backend.Inner.Sources.All(x=>!x.Removed && x.Folder=="trash"),"Movimento automático rápido excluiu definitivamente ou não concluiu.");
        int expensive=backend.ExpensiveCalls;engine.Verify(job,pendingOnly:true,moveValidatedToDeletedItems:true);
        Check(backend.ExpensiveCalls==expensive,"Movimento rápido concluído foi conferido novamente.");
    }
    private static void CopyWithoutMsg(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Copied && x.StageMsgPath is null && x.TargetSnapshotHash is null), "A cópia rápida não ficou copiada com validação profunda pendente.");
        Check(backend.ExpensiveCalls == 0 && backend.NativeCalls == 1 && backend.Inner.Copies.Count == 5, "O caminho rápido executou MSG ou repetiu a cópia.");
        engine.Verify(job, includeFast: false); engine.Remove(job, false);
        Check(backend.ExpensiveCalls == 0 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Validação adiada foi executada ou promoveu estado sem prova.");
        Intact(backend);
        engine.Verify(job);
        Check(backend.ExpensiveCalls > 0 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "A validação explícita não completou as provas de remoção.");
        Intact(backend);
    }
    private static void ReconcileWithoutDeferredMsg(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
            foreach (var row in ledger.GetItems(job))
            { row.State = ArchiveItemState.ReviewRequired; row.Error = "Anexo especial não suportado"; ledger.UpdateItem(row); }
        int copies = backend.NativeCalls;
        engine.Reconcile(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Copied && x.StageMsgPath is not null), "Revisão sem MSG não foi retomada.");
        Check(backend.NativeCalls == copies && backend.Inner.Copies.Count == 5, "Reconciliação da evidência repetiu a cópia.");
        Intact(backend); engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Evidência complementada não passou na validação completa.");
    }
    private static void ReconcileNativeDatePrecision(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
            foreach (var row in ledger.GetItems(job))
            { row.State = ArchiveItemState.ReviewRequired; row.Error = "Received"; ledger.UpdateItem(row); }
        backend.RoundedFallbackIdentity = true;
        backend.Inner.Sources[0].Fields["ReceivedTime"] = FakeArchiveBackend.Hash("date really changed"); backend.Inner.Save();
        int copies = backend.NativeCalls;
        engine.Reconcile(job);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Copied) == 4 &&
            engine.Items(job).Single(x => x.SourceEntryId == "source-0").State == ArchiveItemState.ReviewRequired,
            "Arredondamento bloqueou a leitura nativa ou mudança real de data foi aceita.");
        engine.Verify(job);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 4 && backend.NativeCalls == copies,
            "Recuperação de precisão repetiu a cópia ou não completou a validação.");
        Intact(backend);
    }
    private static void FrozenArrivals(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root));
        backend.Inner.Sources.Add(new() { Id = "late-arrival", RecordKey = "late-record", Size = 32 * 1024, Fields = new(backend.Inner.Sources[0].Fields) }); backend.Inner.Save();
        engine.Copy(job);
        Check(engine.Items(job).Count == 5 && backend.Inner.Copies.Count == 5 && backend.Inner.Copies.All(x => x.SourceId != "late-arrival"), "Nova chegada entrou no escopo congelado.");
        Check(engine.CheckNew(job).NewItemsCount == 1, "Nova chegada não foi informada separadamente."); Intact(backend);
    }
    private static void Resume(string root, string phase)
    {
        string job;
        using (var backend = Provider(root)) using (var engine = new ArchiveEngine(root, backend))
        {
            job = engine.Plan(Config(root));
            engine.Fault = (point, _) => { if (point == phase) throw new ArchiveInjectedCrashException(point); };
            bool crashed = false; try { engine.Copy(job); } catch (ArchiveInjectedCrashException) { crashed = true; }
            Check(crashed, "Ponto de interrupção não atingido: " + phase);
        }
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        next.Copy(job); next.Copy(job);
        Check(resumed.NativeCalls == 1 && resumed.Inner.Copies.Count == 5 && next.Items(job).All(x => x.State == ArchiveItemState.Copied), "Retomada duplicou ou perdeu membros em " + phase);
        Check(resumed.Inner.Copies.Select(x => x.Token).Distinct().Count() == 5 && resumed.Inner.Copies.All(x => x.FolderId == "target-folder"), "Retomada não recuperou tokens/identidades movidas."); Intact(resumed);
    }
    private static void PartialDoesNotRepeat(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); backend.Partial = true; engine.Copy(job); engine.Copy(job);
        Check(backend.NativeCalls == 1 && backend.Inner.Copies.Count == 4, "Resultado parcial foi repetido automaticamente.");
        Check(engine.Items(job).Any(x => x.State == ArchiveItemState.ReviewRequired), "Resultado parcial não ficou pendente."); Intact(backend);
    }
    private static void MetadataChangePending(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); backend.ChangeMetadata = true; engine.Copy(job); engine.Copy(job);
        Check(backend.NativeCalls == 1 && backend.Inner.Copies.Count == 5, "Metadados diferentes causaram recópia automática.");
        Check(engine.Items(job).Any(x => x.State == ArchiveItemState.ReviewRequired), "Metadados diferentes não foram isolados para revisão."); Intact(backend);
    }
    private static void PendingDoesNotBlockRest(string root)
    {
        using var backend = Provider(root);
        for (int i = 5; i < 105; i++) backend.Inner.Sources.Add(new()
        {
            Id = "source-" + i, RecordKey = "record-" + i, Size = 32 * 1024,
            Fields = new(backend.Inner.Sources[0].Fields) { ["Subject"] = FakeArchiveBackend.Hash("later-" + i) }
        });
        backend.Inner.Save(); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); backend.PartialFirstOnly = true; engine.Copy(job); engine.Copy(job);
        Check(backend.NativeCalls == 2 && backend.Inner.Copies.Count == 104, "Lote pendente bloqueou os demais ou foi repetido.");
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Copied) >= 5, "Demais membros do escopo não avançaram após pendência.");
        Intact(backend);
    }
    private static void PartialReconcilePreservesAdopted(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); backend.ChangeMetadata = true; engine.Copy(job);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Copied) == 4, "Correspondências válidas do lote não foram aproveitadas.");
        engine.Verify(job);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 4, "A prova profunda dos membros adotados não foi preservada antes de reconciliar.");
        var unmatched = backend.Inner.Copies.Single(x => x.Token is null);
        var original = backend.Inner.Sources.Single(x => x.Id == unmatched.SourceId);
        int index = backend.Inner.Copies.IndexOf(unmatched);
        backend.Inner.Copies[index] = unmatched with { Fields = new(original.Fields) }; backend.Inner.Save();
        engine.Reconcile(job); engine.Copy(job);
        Check(backend.NativeCalls == 1 && backend.Inner.Copies.Count == 5, "Reconciliação refez a transferência ou duplicou conteúdo.");
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 4 && engine.Items(job).Count(x => x.State == ArchiveItemState.Copied) == 1,
            "Reconciliação não adotou somente a cópia pendente corrigida ou invalidou provas já concluídas.");
        Check(backend.Inner.Copies.Select(x => x.Token).Distinct().Count() == 5, "Reconciliação reutilizou uma correspondência já reivindicada."); Intact(backend);
    }
    private static void DeepBodyDifferenceBlocksRemoval(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); backend.ChangeBody = true; engine.Copy(job);
        Check(backend.ExpensiveCalls == 0 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Conferência rápida indevidamente comparou corpo ou aprovou remoção.");
        engine.Remove(job, false); Intact(backend);
        engine.Verify(job); engine.Remove(job, false);
        Check(engine.Items(job).Any(x => x.State == ArchiveItemState.ReviewRequired), "Diferença real de corpo passou na validação profunda.");
        string changedSourceId = backend.Inner.Copies.First().SourceId;
        Check(backend.Inner.Sources.Single(x => x.Id == changedSourceId).Folder != "trash", "Mensagem com corpo diferente foi removida da origem.");
    }
    private static void CompletedBatchReconcile(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); backend.ChangeBody = true; engine.Copy(job); engine.Verify(job);
        var bad = backend.Inner.Copies.First();
        Check(engine.Items(job).Single(x => x.SourceEntryId == bad.SourceId).State == ArchiveItemState.ReviewRequired, "Pré-condição de revisão profunda não foi atingida.");
        engine.Reconcile(job);
        Check(engine.Items(job).Single(x => x.SourceEntryId == bad.SourceId).State == ArchiveItemState.ReviewRequired && backend.NativeCalls == 1, "Reconciliação aceitou diferença real ou repetiu a transferência.");
        int index = backend.Inner.Copies.IndexOf(bad);
        backend.Inner.Copies[index] = bad with { Fields = new(backend.Inner.Sources.Single(x => x.Id == bad.SourceId).Fields) }; backend.Inner.Save();
        engine.Reconcile(job);
        Check(engine.Items(job).Single(x => x.SourceEntryId == bad.SourceId).State == ArchiveItemState.Copied, "Lote concluído foi ignorado após reavaliação de pendência profunda.");
        Check(backend.NativeCalls == 1 && backend.Inner.Copies.Count == 5, "Reavaliação de lote concluído causou recópia.");
        engine.Verify(job); Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "A reconciliação não permitiu a validação posterior."); Intact(backend);
    }
    private static void FreshVerificationAfterCopy(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        backend.RefreshSourceIdentities(engine.Items(job)); // Simulate a stale provider cache from another operation.
        var changed = backend.Inner.Sources.First(); changed.Fields["Subject"] = FakeArchiveBackend.Hash("changed after copy"); changed.Version++; backend.Inner.Save();
        engine.Verify(job);
        var row = engine.Items(job).Single(x => x.SourceEntryId == changed.Id);
        Check(row.State == ArchiveItemState.ReviewRequired, "Verificação reutilizou identidade em cache e aprovou origem alterada após cópia.");
        Check(changed.Folder != "trash" && !changed.Removed, "Verificação alterou original com identidade divergente.");
    }
    private static void BulkCacheScopeAndCleanup(string root)
    {
        using var backend = Provider(root);
        for (int i = 5; i < 105; i++) backend.Inner.Sources.Add(new()
        {
            Id = "source-" + i, RecordKey = "record-" + i, Size = 32 * 1024,
            Fields = new(backend.Inner.Sources[0].Fields) { ["Subject"] = FakeArchiveBackend.Hash("bulk-" + i) }
        });
        backend.Inner.Save(); using var engine = new ArchiveEngine(root, backend); string job = engine.Plan(Config(root));
        backend.Inner.Sources.Add(new() { Id = "outside-scope", RecordKey = "outside-record", Size = 32 * 1024, Fields = new(backend.Inner.Sources[0].Fields) }); backend.Inner.Save();
        engine.Fault = (phase, _) => { if (phase == "after-batch-copy-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
        bool crashed = false; try { engine.Copy(job); } catch (ArchiveInjectedCrashException) { crashed = true; }
        Check(crashed && backend.RefreshCalls.Count > 0 && backend.RefreshCalls.All(x => x.Count <= 100 && !x.Contains("outside-scope")), "Consulta em lote ultrapassou limite ou escopo congelado.");
        Check(backend.CacheEntries == 0 && backend.ClearCalls > 0, "Interrupção deixou identidades em cache para uma retomada posterior.");
        engine.Fault = null; engine.Copy(job);
        Check(backend.NativeCalls == 2 && backend.Inner.Copies.Count == 105 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Retomada com leitura em lote duplicou ou perdeu mensagens.");
        Check(backend.TargetRefreshCalls > 0 && backend.CacheEntries == 0, "Comparação do destino não usou consulta em lote ou manteve cache ao concluir."); Intact(backend);
    }

    private static void BulkFinalizeResume(string root)
    {
        using var backend = new FinalizeFake(Path.Combine(root, "provider.json")); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root));
        engine.Fault = (phase, _) => { if (phase == "after-batch-finalize-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
        bool crashed = false; try { engine.Copy(job); } catch (ArchiveInjectedCrashException) { crashed = true; }
        Check(crashed && backend.FinalizeCalls == 1 && backend.Inner.Copies.All(x => x.FolderId == "target-folder"), "Interrupção não ocorreu após a finalização em lote.");
        engine.Fault = null; engine.Copy(job);
        Check(backend.NativeCalls == 1 && backend.FinalizeCalls == 2 && backend.Inner.Copies.Count == 5 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Finalização retomada duplicou mensagens ou repetiu a cópia original.");
        engine.Remove(job, false); Intact(backend);
    }
    private static void BulkFinalizePartial(string root)
    {
        using var backend = new FinalizeFake(Path.Combine(root, "provider.json")) { InterruptAfterTwo = true }; using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(backend.Inner.Copies.Count(x => x.FolderId == "target-folder") == 2 && engine.Items(job).All(x => x.State == ArchiveItemState.ReviewRequired), "Movimentação interna parcial não reteve revisão.");
        backend.InterruptAfterTwo = false; engine.Reconcile(job);
        Check(backend.NativeCalls == 1 && backend.Inner.Copies.Count == 5 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Reconciliação interna parcial não preservou as cópias únicas."); Intact(backend);
    }
    private sealed class FinalizeFake(string path) : FastFake(path), IArchiveBulkFinalizeBackend
    {
        internal int FinalizeCalls;
        internal bool InterruptAfterTwo;
        public Dictionary<string, DestinationIdentity> FinalizeBatchCopies(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> rows)
        {
            FinalizeCalls++; var result = new Dictionary<string, DestinationIdentity>();
            foreach (var row in rows)
            {
                MarkBatchCopy(volume, batch, row, batch.Matches[row.Key]);
                result[row.Key] = FinalizeBatchCopy(volume, batch, row, batch.Matches[row.Key]);
                if (InterruptAfterTwo && result.Count == 2) throw new ArchiveReviewException("Finalização artificial parcial; reconcilie sem recópia.");
            }
            return result;
        }
    }
    private class FastFake : IArchiveBackend, IArchiveBatchBackend, IArchiveBulkFastBatchBackend
    {
        internal readonly FakeArchiveBackend Inner;
        private readonly string callsPath;
        internal int NativeCalls, ExpensiveCalls;
        internal int ClearCalls, TargetRefreshCalls;
        internal List<List<string>> RefreshCalls = [];
        private readonly Dictionary<string, FastMessageIdentity> sourceCache = new(StringComparer.Ordinal);
        private readonly Dictionary<string, FastMessageIdentity> targetCache = new(StringComparer.Ordinal);
        internal int CacheEntries => sourceCache.Count + targetCache.Count;
        internal bool Partial, PartialFirstOnly, ChangeMetadata, ChangeBody, RoundedFallbackIdentity;
        internal FastFake(string path)
        {
            Inner = new(path); callsPath = path + ".calls";
            if (File.Exists(callsPath)) NativeCalls = int.Parse(File.ReadAllText(callsPath));
            else { foreach (var source in Inner.Sources) source.Size = 32 * 1024; Inner.Save(); }
        }
        private static Dictionary<string, string> Light(Dictionary<string, string> fields) => fields.Where(x => !x.Key.StartsWith("NativeBody-", StringComparison.Ordinal)).ToDictionary(x => x.Key, x => x.Value);
        public FastMessageIdentity CaptureSourceIdentity(ArchiveItem item)
        {
            if(UnavailableNative)throw new ArchiveReviewException("Identificação nativa da mensagem pendente: message unavailable");
            if (sourceCache.TryGetValue(item.SourceEntryId, out var cached)) return cached;
            var source = Inner.Sources.Single(x => x.Id == item.SourceEntryId && !x.Removed && x.Folder == item.SourceFolderId);
            var fields = Light(source.Fields);
            if (RoundedFallbackIdentity) fields["ReceivedTime"] = FakeArchiveBackend.Hash("rounded OOM value");
            return new(fields, source.Version.ToString());
        }
        public FastMessageIdentity CaptureTargetIdentity(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity) => targetCache.TryGetValue(identity.EntryId, out var cached) ? cached : new(Light(Inner.Copies.Single(x => x.EntryId == identity.EntryId).Fields));
        public void RefreshSourceIdentities(IReadOnlyList<ArchiveItem> rows)
        {
            if(UnavailableNative && rows.Count>0)throw new ArchiveReviewException("Identificação nativa da mensagem pendente: message unavailable");
            RefreshCalls.Add(rows.Select(x => x.SourceEntryId).ToList());
            foreach (var row in rows)
            {
                var source = Inner.Sources.Single(x => x.Id == row.SourceEntryId && !x.Removed && x.Folder == row.SourceFolderId);
                sourceCache[row.SourceEntryId] = new(Light(source.Fields), source.Version.ToString());
            }
        }
        public void RefreshTargetIdentities(ArchiveVolume volume, IReadOnlyList<DestinationIdentity> candidates)
        {
            TargetRefreshCalls++;
            foreach (var identity in candidates) targetCache[identity.EntryId] = new(Light(Inner.Copies.Single(x => x.EntryId == identity.EntryId).Fields));
        }
        public void ClearFastIdentityCache() { ClearCalls++; sourceCache.Clear(); targetCache.Clear(); }
        internal bool UnavailableNative;
        public bool CanBatch(ArchiveItem item) => true;
        public BatchFolder PrepareBatch(ArchiveVolume volume, ArchiveBatch batch) { Inner.AttachVolume(volume); return new(volume.StoreId!, "batch-" + batch.Id); }
        public void AssertBatchIdle(ArchiveBatch batch) { }
        public void CopyBatch(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> items)
        {
            NativeCalls++; File.WriteAllText(callsPath, NativeCalls.ToString());
            foreach (var item in Partial || PartialFirstOnly && NativeCalls == 1 ? items.Take(items.Count - 1) : items)
            {
                var source = Inner.Sources.Single(x => x.Id == item.SourceEntryId); var fields = new Dictionary<string, string>(source.Fields);
                if (Inner.Copies.Count == 0 && ChangeMetadata) fields["Subject"] = FakeArchiveBackend.Hash("unexpected subject");
                if (Inner.Copies.Count == 0 && ChangeBody) fields["NativeBody-MSGBodyStream:__substg1.0_1000001F"] = FakeArchiveBackend.Hash("unexpected body");
                Inner.Copies.Add(new(volume.JobId, volume.Number, source.Id, null, Guid.NewGuid().ToString("N"), batch.FolderId!, fields, source.Size));
                using var stream = new FileStream(volume.PstPath, FileMode.Open, FileAccess.Write); stream.SetLength(stream.Length + source.Size);
            }
            Inner.Save();
        }
        public List<DestinationIdentity> BatchCandidates(ArchiveVolume volume, ArchiveBatch batch) => Inner.Copies.Where(x => x.JobId == volume.JobId && x.FolderId == batch.FolderId).Select(x => new DestinationIdentity(volume.StoreId!, x.FolderId, x.EntryId)).ToList();
        public void MarkBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem item, DestinationIdentity identity)
        {
            var copies = Inner.Copies.Where(x => x.EntryId == identity.EntryId || x.Token == item.Token).ToList();
            if (copies.Count != 1) throw new ArchiveReviewException("Identificação artificial ambígua.");
            int index = Inner.Copies.IndexOf(copies[0]); Inner.Copies[index] = copies[0] with { Token = item.Token }; Inner.Save();
        }
        public DestinationIdentity FinalizeBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem item, DestinationIdentity identity)
        {
            var found = Inner.Copies.Single(x => x.Token == item.Token); int index = Inner.Copies.IndexOf(found);
            if (found.FolderId != "target-folder") { found = found with { FolderId = "target-folder", EntryId = Guid.NewGuid().ToString("N") }; Inner.Copies[index] = found; Inner.Save(); }
            return new(volume.StoreId!, found.FolderId, found.EntryId);
        }
        public void CloseBatch(ArchiveVolume volume, ArchiveBatch batch) { }
        public List<SourceDescriptor> Discover(ArchiveConfig config) => Inner.Discover(config);
        public List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config) => Inner.CaptureScope(config);
        public List<SourceDescriptor> DiscoverFrozen(ArchiveConfig config, IReadOnlyList<FolderScopeSnapshot> scope) => Inner.DiscoverFrozen(config, scope);
        public ArchiveSnapshot Export(ArchiveItem item, string directory) { ExpensiveCalls++; return Inner.Export(item, directory); }
        public ArchiveSnapshot SourceSnapshot(ArchiveItem item, SourceLocation? location, string directory) { ExpensiveCalls++; return Inner.SourceSnapshot(item, location, directory); }
        public bool SourceMatchesChangeKey(ArchiveItem item, string expected) => Inner.SourceMatchesChangeKey(item, expected);
        public ArchiveSnapshot BackupSnapshot(ArchiveItem item, string directory) { ExpensiveCalls++; return Inner.BackupSnapshot(item, directory); }
        public void CreateVolume(ArchiveVolume volume) => Inner.CreateVolume(volume);
        public void AttachVolume(ArchiveVolume volume) => Inner.AttachVolume(volume);
        public long MeasureVolume(ArchiveVolume volume) => Inner.MeasureVolume(volume);
        public TokenLookup FindToken(ArchiveVolume volume, ArchiveItem item) => Inner.FindToken(volume, item);
        public DestinationIdentity Import(ArchiveVolume volume, ArchiveItem item) => throw new Exception("O fluxo rápido caiu no importador MSG.");
        public ArchiveSnapshot TargetSnapshot(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity, string directory, bool requireToken = true) { ExpensiveCalls++; return Inner.TargetSnapshot(volume, item, identity, directory, requireToken); }
        public List<DestinationIdentity> UnmarkedCandidates(ArchiveVolume volume, ArchiveItem item) => Inner.UnmarkedCandidates(volume, item);
        public void MarkCopy(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity) => Inner.MarkCopy(volume, item, identity);
        public DiskRoundtrip ReopenVolume(ArchiveVolume volume) => Inner.ReopenVolume(volume);
        public SourceLocation? LocateSource(ArchiveItem item) => Inner.LocateSource(item);
        public SourceLocation MoveToDeletedItems(ArchiveItem item, string? expected) => Inner.MoveToDeletedItems(item, expected);
        public void DeleteFromDeletedItems(ArchiveItem item, SourceLocation location, string? expected) => Inner.DeleteFromDeletedItems(item, location, expected);
        public void Dispose() => Inner.Dispose();
    }
}

