using System.Text.Json;
using Microsoft.Data.Sqlite;

// Exercises the actual catalog and engine against a durable, synthetic provider.
// No COM, Outlook profile, real mailbox or user catalog is used by this suite.
internal static class ArchiveBatchTests
{
    internal static int Run()
    {
        string root = Path.Combine(Path.GetTempPath(), "archive-batch-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var tests = new (string Name, Action<string> Test)[]
        {
            ("complete-and-verify", CompleteAndVerify),
            ("repair-history", LaterRepairPreservesCompletedBatch),
            ("scoped-reconciliation", ScopedReconciliation),
            ("identical-content-multiplicity", IdenticalContentMultiplicity),
            ("resume-native-copy", path => ResumeEffect(path, Effect.Copy)),
            ("resume-token-save", path => ResumeEffect(path, Effect.Mark)),
            ("resume-final-move", path => ResumeEffect(path, Effect.Finalize)),
            ("resume-durable-matches", ResumeDurableMatches),
            ("empty-ambiguous-outcome", EmptyOutcomeNeverRetries),
            ("partial-outcome", path => InvalidOutcome(path, Outcome.Partial)),
            ("extra-candidate", path => InvalidOutcome(path, Outcome.Extra)),
            ("changed-candidate", path => InvalidOutcome(path, Outcome.Changed)),
            ("worker-still-active", WorkerStillActive),
            ("worker-active-mixed", WorkerBlocksRemaining),
            ("backup-changed-before-recovery", BackupChangedBeforeRecovery),
            ("legacy-intent-compatible", LegacyIntentCompatible),
            ("legacy-fallback", LegacyFallback),
            ("volume-separation", VolumeSeparation),
            ("folder-separation", FolderSeparation),
            ("batch-count-limit", BatchCountLimit),
            ("equal-content-across-jobs", EqualContentAcrossJobs),
            ("writer-version-guard", WriterVersionGuard)
        };
        try
        {
            foreach (var test in tests)
            {
                string scenario = Path.Combine(root, test.Name); Directory.CreateDirectory(scenario);
                test.Test(scenario);
                Console.WriteLine("Lotes: " + test.Name + " aprovado.");
            }
            Console.WriteLine($"Lotes: {tests.Length} cenários de retomada, correspondência, limites e compatibilidade aprovados.");
            return 0;
        }
        finally
        {
            // Remove only this suite's newly created synthetic temporary tree.
            string full = Path.GetFullPath(root);
            string temp = Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (full.StartsWith(temp, StringComparison.OrdinalIgnoreCase) && Path.GetFileName(full).StartsWith("archive-batch-tests-", StringComparison.Ordinal))
                Directory.Delete(full, true);
        }
    }

    private static ArchiveConfig Config(string root, long target = 128 * ArchiveVolumePolicy.MiB) => new()
    {
        RootDirectory = root, SourceStoreId = "source", SourceFolderId = "folder", IncludeSubfolders = true,
        TargetVolumeBytes = target, HardLimitBytes = 256 * ArchiveVolumePolicy.MiB, ReserveBytes = ArchiveVolumePolicy.MiB
    };
    private static BatchFake Provider(string root) => new(Path.Combine(root, "provider.json"));
    private static void ScopedReconciliation(string root)
    {
        using var backend=Provider(root);using var engine=new ArchiveEngine(root,backend);
        string job=engine.Plan(Config(root));backend.CopyOutcome=Outcome.Extra;engine.Copy(job);
        var keys=engine.Items(job).Select(x=>x.Key).ToHashSet(StringComparer.Ordinal);
        int Version(){using var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);return ledger.GetBatches(job).Single().Version;}
        int before=Version(),copies=backend.Inner.Copies.Count,calls=backend.NativeCalls;
        engine.Reconcile(job,new HashSet<string>{keys.First()});
        Check(Version()==before,"Reconciliação parcial alterou um lote nativo fora do escopo completo.");
        engine.Reconcile(job,keys);
        Check(Version()>before,"Reconciliação por lote não revisitou o resultado nativo pendente.");
        Check(backend.NativeCalls==calls && backend.Inner.Copies.Count==copies,"Reconciliação repetiu a transferência nativa.");
        SourcesIntact(backend);
    }
    private static void Check(bool condition, string message)
    { if (!condition) throw new Exception("Lotes: " + message); }
    private static List<ArchiveItem> ReadItems(string root, string job)
    { using var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite"), true); return ledger.GetItems(job); }
    private static void ExpectCrash(Action action, string point)
    {
        try { action(); }
        catch (ArchiveInjectedCrashException) { return; }
        throw new Exception("Lotes: a interrupção artificial não ocorreu em " + point);
    }
    private static void SourcesIntact(BatchFake backend)
    { Check(backend.Inner.Sources.All(x => !x.Removed && x.Folder != "trash") && backend.Inner.DeleteCalls == 0, "O ensaio de cópia alterou ou removeu originais."); }

    private static void LaterRepairPreservesCompletedBatch(string root)
    {
        using var backend=Provider(root);backend.Inner.EmitRelevantMetadata=true;backend.CopyOutcome=Outcome.Partial;
        using var engine=new ArchiveEngine(root,backend);string job=engine.Plan(Config(root));engine.Copy(job);engine.Verify(job);
        var selected=engine.Items(job).Single(x=>x.State==ArchiveItemState.ReviewRequired);
        var source=backend.Inner.Sources.Single(x=>x.Id==selected.SourceEntryId);source.Fields["Subject"]=FakeArchiveBackend.Hash("origem atualizada");source.Version++;backend.Inner.Save();
        engine.ResolvePending(job,new HashSet<string>{selected.Key},true);engine.Verify(job,pendingOnly:true);engine.Remove(job,false);
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
            foreach(var row in ledger.GetItems(job)){row.State=ArchiveItemState.ReviewRequired;ledger.UpdateItem(row);}
        int calls=backend.NativeCalls;engine.Reconcile(job);
        Check(backend.NativeCalls==calls && engine.Items(job).All(x=>x.State==ArchiveItemState.InDeletedItems),"Reparo posterior invalidou o lote anterior ou repetiu a transferência.");
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))Check(ledger.GetBatches(job).Single().State==ArchiveBatchState.Complete,"Lote finalizado não foi reconhecido.");
    }
    private static void CompleteAndVerify(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(engine.Items(job).Count == 5 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "O lote completo não foi registrado como cinco cópias.");
        Check(backend.NativeCalls == 1 && backend.LegacyCalls == 0 && backend.Inner.Copies.Count == 5, "A cópia completa não usou exatamente um lote sem fallback.");
        Check(backend.Inner.Copies.All(x => x.FolderId == "target-folder" && !string.IsNullOrEmpty(x.Token)), "As cópias não foram identificadas/finalizadas.");
        engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "A verificação completa posterior ao lote não passou.");
        SourcesIntact(backend);
    }
    private static void IdenticalContentMultiplicity(string root)
    {
        using var backend = Provider(root);
        foreach (var source in backend.Inner.Sources.Skip(1)) source.Fields = new(backend.Inner.Sources[0].Fields);
        backend.Inner.Save(); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Mensagens iguais perderam a correspondência um a um.");
        Check(backend.Inner.Copies.Count == 5 && backend.Inner.Copies.Select(x => x.EntryId).Distinct().Count() == 5 && backend.Inner.Copies.Select(x => x.Token).Distinct().Count() == 5,
            "Conteúdos iguais foram colapsados ou compartilharam uma cópia/token.");
        SourcesIntact(backend);
    }
    private enum Effect { Copy, Mark, Finalize }
    private static void ResumeEffect(string root, Effect effect)
    {
        string job;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            job = engine.Plan(Config(root));
            backend.CrashAfter = effect;
            ExpectCrash(() => engine.Copy(job), effect.ToString());
        }
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            engine.Copy(job); engine.Verify(job);
            Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Retomada não concluiu após efeito " + effect);
            Check(backend.NativeCalls == 1 && backend.LegacyCalls == 0 && backend.Inner.Copies.Count == 5, "Retomada repetiu uma chamada nativa ou criou duplicatas após " + effect);
            Check(backend.Inner.Copies.All(x => x.FolderId == "target-folder"), "Retomada não reconciliou a pasta final.");
            SourcesIntact(backend);
        }
    }
    private static void ResumeDurableMatches(string root)
    {
        string job; bool injected = false;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            job = engine.Plan(Config(root));
            engine.Fault = (point, _) => { if (point == "after-batch-matched") { injected = true; throw new ArchiveInjectedCrashException(point); } };
            ExpectCrash(() => engine.Copy(job), "after-batch-matched");
        }
        Check(injected, "O mapa durável não possui ponto de interrupção verificável.");
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        next.Copy(job); next.Verify(job);
        Check(next.Items(job).All(x => x.State == ArchiveItemState.Verified) && resumed.NativeCalls == 1 && resumed.Inner.Copies.Count == 5,
            "O mapa durável não foi reutilizado sem repetir o lote.");
    }
    private static void EmptyOutcomeNeverRetries(string root)
    {
        string job;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            job = engine.Plan(Config(root)); backend.CrashBeforeCopy = true;
            ExpectCrash(() => engine.Copy(job), "before-copy-effect");
        }
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        next.Copy(job); next.Copy(job);
        Check(resumed.NativeCalls == 1 && resumed.Inner.Copies.Count == 0 && next.Items(job).All(x => x.State == ArchiveItemState.ReviewRequired),
            "Resultado vazio depois da intenção ambígua foi repetido automaticamente.");
        SourcesIntact(resumed);
    }
    private enum Outcome { Complete, Partial, Extra, Changed }
    private static void InvalidOutcome(string root, Outcome outcome)
    {
        string job;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            job = engine.Plan(Config(root)); backend.CopyOutcome = outcome; engine.Copy(job);
        }
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        int existing = resumed.Inner.Copies.Count; next.Copy(job); next.Verify(job);
        Check(resumed.NativeCalls == 1 && resumed.LegacyCalls == 0 && resumed.Inner.Copies.Count == existing, "Resultado inválido foi repetido ou recebeu fallback: " + outcome);
        if (outcome == Outcome.Extra)
        {
            Check(next.Items(job).All(x => x.State == ArchiveItemState.ReviewRequired), "Resultado com cópias extras não bloqueou todo o lote.");
            Check(resumed.Inner.Copies.All(x => x.Token is null), "Candidatos com contagem excedente foram adotados.");
        }
        else
        {
            Check(next.Items(job).Count(x => x.State == ArchiveItemState.ReviewRequired) == 1 && next.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 4,
                "Resultado parcial/alterado não isolou somente a mensagem sem correspondência: " + outcome);
            Check(resumed.Inner.Copies.Count(x => x.Token is not null) == 4 && resumed.Inner.Copies.Where(x => x.Token is not null).All(x => x.FolderId == "target-folder"),
                "Correspondências integrais do resultado parcial/alterado não foram preservadas: " + outcome);
        }
        SourcesIntact(resumed);
    }
    private static void WorkerStillActive(string root)
    {
        string job;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        { job = engine.Plan(Config(root)); backend.CrashAfter = Effect.Copy; ExpectCrash(() => engine.Copy(job), "copy"); }
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            backend.WorkerActive = true;
            try { engine.Copy(job); } catch (IOException) { } catch (ArchiveReviewException) { }
            Check(backend.NativeCalls == 1 && backend.Inner.Copies.Count == 5 && backend.Inner.Copies.All(x => x.Token is null), "O motor reconciliou/repetiu um lote com worker ativo.");
            SourcesIntact(backend);
        }
    }
    private static void BackupChangedBeforeRecovery(string root)
    {
        string job;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        { job = engine.Plan(Config(root)); backend.CrashAfter = Effect.Copy; ExpectCrash(() => engine.Copy(job), "copy"); }
        var row = ReadItems(root, job).First(); File.AppendAllText(row.StageMsgPath!, "corruption");
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        try { next.Copy(job); } catch (ArchiveReviewException) { }
        Check(resumed.NativeCalls == 1 && resumed.Inner.Copies.Count == 5 && !next.Items(job).Any(x => x.State == ArchiveItemState.Verified),
            "Backup alterado permitiu repetir ou verificar o lote.");
        SourcesIntact(resumed);
    }
    private static void WorkerBlocksRemaining(string root)
    {
        string job;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            for (int i = 5; i < 55; i++) backend.Inner.Sources.Add(new FakeArchiveBackend.FakeSource
            { Id = "source-" + i, RecordKey = "record-" + i, Size = 32 * 1024,
                Fields = new() { ["Subject"] = FakeArchiveBackend.Hash("artificial-" + i), ["Body"] = FakeArchiveBackend.Hash("body-" + i) } });
            backend.Inner.Save(); job = engine.Plan(Config(root)); backend.CrashAfter = Effect.Copy;
            ExpectCrash(() => engine.Copy(job), "copy-first-fifty");
        }
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        resumed.WorkerActive = true; next.Copy(job); next.Verify(job); next.Remove(job, false);
        Check(resumed.NativeCalls == 1 && resumed.LegacyCalls == 0 && resumed.Inner.Copies.Count == 50 && resumed.ReopenCalls == 0,
            "Worker ativo permitiu copiar outros itens, reabrir o PST ou repetir o lote.");
        SourcesIntact(resumed);
    }
    private static void LegacyIntentCompatible(string root)
    {
        string job, legacyKey;
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        {
            backend.AllEligible = false; job = engine.Plan(Config(root)); legacyKey = engine.Items(job).First().Key;
            engine.Fault = (point, _) => { if (point == "after-import-before-checkpoint") throw new ArchiveInjectedCrashException(point); };
            ExpectCrash(() => engine.Copy(job), "legacy-import");
        }
        using var resumed = Provider(root); using var next = new ArchiveEngine(root, resumed);
        next.Copy(job); next.Verify(job);
        Check(next.Items(job).All(x => x.State == ArchiveItemState.Verified) && resumed.Inner.Copies.Count == 5 && resumed.LegacyCalls == 1,
            "A intenção individual antiga foi reinterpretada ou duplicada na retomada.");
        Check(next.Items(job).Single(x => x.Key == legacyKey).NativeBatchId is null, "Uma intenção individual antiga entrou em lote nativo.");
        Check(resumed.NativeCalls == 1, "Os novos itens não puderam usar lote ao lado da intenção antiga.");
        SourcesIntact(resumed);
    }
    private static void LegacyFallback(string root)
    {
        using var backend = Provider(root); backend.AllEligible = false;
        using var engine = new ArchiveEngine(root, backend); string job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
        Check(backend.NativeCalls == 0 && backend.LegacyCalls == 5 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "Mensagens não elegíveis não mantiveram o transporte individual.");
    }
    private static void VolumeSeparation(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string job = engine.Plan(Config(root, 3 * ArchiveVolumePolicy.MiB)); engine.Copy(job); engine.Verify(job);
        using var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite"), true);
        Check(ledger.GetVolumes(job).Count >= 3 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "A rotação de volumes não foi respeitada pelo lote.");
        Check(backend.Calls.All(call => call.Keys.Select(key => engine.Items(job).Single(x => x.Key == key).VolumeNumber).Distinct().Count() == 1), "Um lote cruzou volumes.");
    }
    private static void FolderSeparation(string root)
    {
        using var backend = Provider(root); backend.Inner.Sources[3].Folder = "folder-two"; backend.Inner.Save();
        using var engine = new ArchiveEngine(root, backend); string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Copied) && backend.Calls.Count >= 2,
            "Pastas distintas não foram copiadas em lotes separados.");
        Check(backend.Calls.All(call => call.Keys.Select(key => engine.Items(job).Single(x => x.Key == key).SourceFolderId).Distinct().Count() == 1), "Um lote cruzou pastas de origem.");
    }
    private static void BatchCountLimit(string root)
    {
        using var backend = Provider(root);
        for (int i = 5; i < 55; i++) backend.Inner.Sources.Add(new FakeArchiveBackend.FakeSource
        {
            Id = "source-" + i, RecordKey = "record-" + i, Size = 32 * 1024,
            Fields = new() { ["Subject"] = FakeArchiveBackend.Hash("artificial-" + i), ["Body"] = FakeArchiveBackend.Hash("body-" + i) }
        });
        backend.Inner.Save(); using var engine = new ArchiveEngine(root, backend); string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(engine.Items(job).Count == 55 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "O conjunto de 55 itens não foi copiado.");
        Check(backend.Calls.Count >= 2 && backend.Calls.All(x => x.Keys.Count <= 50) && backend.Calls.Sum(x => x.Keys.Count) == 55,
            "O limite de 50 itens por chamada não foi mantido.");
    }
    private static void EqualContentAcrossJobs(string root)
    {
        using var backend = Provider(root); using var engine = new ArchiveEngine(root, backend);
        string first = engine.Plan(Config(root)); engine.Copy(first);
        backend.Inner.Sources.Add(new FakeArchiveBackend.FakeSource { Id = "source-new-identical", RecordKey = "record-new", Size = 32 * 1024, Fields = new(backend.Inner.Sources[0].Fields) });
        backend.Inner.Save(); string second = engine.Plan(Config(root)); engine.Copy(second);
        Check(engine.Items(second).Count == 1 && engine.Items(second).Single().State == ArchiveItemState.Copied && backend.Inner.Copies.Count == 6,
            "Um novo original de conteúdo igual foi colapsado contra outro lote/execução.");
    }
    private static void WriterVersionGuard(string root)
    {
        using (var backend = Provider(root))
        using (var engine = new ArchiveEngine(root, backend))
        { string job = engine.Plan(Config(root)); engine.Copy(job); }
        using var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = Path.Combine(root, "archive.sqlite"), Mode = SqliteOpenMode.ReadOnly, Pooling = false }.ToString());
        connection.Open(); using var command = connection.CreateCommand(); command.CommandText = "PRAGMA application_id";
        Check(Convert.ToInt64(command.ExecuteScalar()) != 1128355889, "O catálogo com lotes ainda é aceito por um escritor antigo.");
    }

    private sealed record Call(string BatchId, int Volume, string Folder, List<string> Keys);
    private sealed class CounterState
    {
        public int NativeCalls { get; set; }
        public int LegacyCalls { get; set; }
        public List<Call> Calls { get; set; } = [];
    }
    private sealed class BatchFake : IArchiveBackend, IArchiveBatchBackend
    {
        internal readonly FakeArchiveBackend Inner;
        private readonly string countersPath;
        private readonly CounterState counters;
        internal bool AllEligible = true, WorkerActive, CrashBeforeCopy;
        internal Effect? CrashAfter;
        internal Outcome CopyOutcome;
        internal int NativeCalls => counters.NativeCalls;
        internal int LegacyCalls => counters.LegacyCalls;
        internal int ReopenCalls { get; private set; }
        internal List<Call> Calls => counters.Calls;
        internal BatchFake(string file)
        {
            Inner = new FakeArchiveBackend(file); countersPath = file + ".native.json";
            counters = File.Exists(countersPath) ? JsonSerializer.Deserialize<CounterState>(File.ReadAllText(countersPath))! : new();
            if (!File.Exists(countersPath)) { foreach (var source in Inner.Sources) source.Size = 32 * 1024; Inner.Save(); SaveCounters(); }
        }
        private void SaveCounters() => File.WriteAllText(countersPath, JsonSerializer.Serialize(counters));
        private void Crash(Effect effect)
        { if (CrashAfter == effect) { CrashAfter = null; throw new ArchiveInjectedCrashException("native-effect-" + effect); } }
        public bool CanBatch(ArchiveItem item) => AllEligible;
        public BatchFolder PrepareBatch(ArchiveVolume volume, ArchiveBatch batch)
        { Inner.AttachVolume(volume); return new(volume.StoreId!, "batch-folder-" + batch.Id); }
        public void AssertBatchIdle(ArchiveBatch batch)
        { if (WorkerActive) throw new IOException("Worker artificial ainda ativo."); }
        public void CopyBatch(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> items)
        {
            counters.NativeCalls++; counters.Calls.Add(new(batch.Id, volume.Number, batch.SourceFolderId, items.Select(x => x.Key).ToList())); SaveCounters();
            if (CrashBeforeCopy) throw new ArchiveInjectedCrashException("native-before-effect");
            var selected = CopyOutcome == Outcome.Partial ? items.Take(Math.Max(0, items.Count - 1)) : items;
            foreach (var item in selected)
            {
                var source = Inner.Sources.Single(x => x.Id == item.SourceEntryId);
                var fields = new Dictionary<string, string>(source.Fields);
                if (CopyOutcome == Outcome.Changed && Inner.Copies.All(x => x.FolderId != batch.FolderId)) fields["Subject"] = FakeArchiveBackend.Hash("changed by provider");
                Add(volume, batch.FolderId!, source.Id, fields, source.Size);
            }
            if (CopyOutcome == Outcome.Extra)
            {
                var source = Inner.Sources.First(); Add(volume, batch.FolderId!, source.Id, new(source.Fields), source.Size);
            }
            Inner.Save(); Crash(Effect.Copy);
        }
        private void Add(ArchiveVolume volume, string folder, string source, Dictionary<string, string> fields, long size)
        {
            Inner.Copies.Add(new FakeArchiveBackend.FakeCopy(volume.JobId, volume.Number, source, null, Guid.NewGuid().ToString("N"), folder, fields, size));
            using var file = new FileStream(volume.PstPath, FileMode.Open, FileAccess.Write); file.SetLength(file.Length + size);
        }
        public List<DestinationIdentity> BatchCandidates(ArchiveVolume volume, ArchiveBatch batch)
        { AssertBatchIdle(batch); return Inner.Copies.Where(x => x.JobId == volume.JobId && x.Volume == volume.Number && x.FolderId == batch.FolderId).Select(x => new DestinationIdentity(volume.StoreId!, x.FolderId, x.EntryId)).ToList(); }
        public void MarkBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem item, DestinationIdentity identity)
        {
            int index = Inner.Copies.FindIndex(x => x.JobId == volume.JobId && x.EntryId == identity.EntryId);
            if (index < 0)
            {
                var moved = Inner.Copies.Where(x => x.JobId == volume.JobId && x.Token == item.Token && x.FolderId == "target-folder").ToList();
                if (moved.Count == 1) return; // Durable token reconciles an already-moved copy whose EntryID changed.
            }
            if (index < 0 || Inner.Copies[index].FolderId != batch.FolderId) throw new ArchiveReviewException("Candidato fora do lote.");
            var copy = Inner.Copies[index];
            if (copy.Token is not null && copy.Token != item.Token) throw new ArchiveReviewException("Token estranho.");
            Inner.Copies[index] = copy with { Token = item.Token }; Inner.Save(); Crash(Effect.Mark);
        }
        public DestinationIdentity FinalizeBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem item, DestinationIdentity identity)
        {
            var matches = Inner.Copies.Where(x => x.JobId == volume.JobId && x.Token == item.Token && (x.FolderId == batch.FolderId || x.FolderId == "target-folder")).ToList();
            if (matches.Count != 1) throw new ArchiveReviewException("Token ambíguo durante finalização.");
            var found = matches[0];
            if (found.FolderId != "target-folder")
            {
                int index = Inner.Copies.IndexOf(found); found = found with { FolderId = "target-folder", EntryId = Guid.NewGuid().ToString("N") };
                Inner.Copies[index] = found; Inner.Save(); Crash(Effect.Finalize);
            }
            return new(volume.StoreId!, found.FolderId, found.EntryId);
        }
        public void CloseBatch(ArchiveVolume volume, ArchiveBatch batch)
        { Check(!Inner.Copies.Any(x => x.JobId == volume.JobId && x.FolderId == batch.FolderId), "Uma pasta de lote ainda contém cópias ao ser concluída."); }
        public List<SourceDescriptor> Discover(ArchiveConfig config) => Inner.Discover(config);
        public List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config) => Inner.CaptureScope(config);
        public List<SourceDescriptor> DiscoverFrozen(ArchiveConfig config, IReadOnlyList<FolderScopeSnapshot> scope) => Inner.DiscoverFrozen(config, scope);
        private static ArchiveSnapshot Profile(ArchiveSnapshot snapshot, string directory)
        {
            var fields = new Dictionary<string, string>
            {
                ["ProfileVersion"] = "relevant-message-v1", ["BodyMode"] = "html-relevant-v1",
                ["Body-Text"] = ArchiveFiles.HashFields(snapshot.Fields.Where(x => x.Key.StartsWith("NativeBody-", StringComparison.Ordinal) || x.Key == "Body").ToDictionary(x => x.Key, x => x.Value))
            };
            foreach (string name in new[] { "Links", "Images", "Resources", "Tables", "Lists", "Preformatted", "ActivePayloads" }) fields["Body-" + name] = FakeArchiveBackend.Hash("[]");
            foreach (string name in new[] { "SenderName", "SenderEmailType", "SenderEmailAddress", "SentOnBehalfOfName", "SentUtc", "ReceivedUtc", "Recipients" })
                fields["Meta-" + name] = snapshot.Fields.GetValueOrDefault(name, FakeArchiveBackend.Hash("synthetic-" + name));
            string sidecar = Path.Combine(directory, "relevant-fields.json");
            if (!File.Exists(sidecar)) ArchiveFiles.WriteNew(sidecar, fields);
            return snapshot with { RelevantFields = fields };
        }
        public ArchiveSnapshot Export(ArchiveItem item, string directory) => Profile(Inner.Export(item, directory), directory);
        public ArchiveSnapshot SourceSnapshot(ArchiveItem item, SourceLocation? location, string directory) => Profile(Inner.SourceSnapshot(item, location, directory), directory);
        public bool SourceMatchesChangeKey(ArchiveItem item, string expected) => Inner.SourceMatchesChangeKey(item, expected);
        public ArchiveSnapshot BackupSnapshot(ArchiveItem item, string directory) => Profile(Inner.BackupSnapshot(item, directory), directory);
        public void CreateVolume(ArchiveVolume volume) => Inner.CreateVolume(volume);
        public void AttachVolume(ArchiveVolume volume) => Inner.AttachVolume(volume);
        public long MeasureVolume(ArchiveVolume volume) => Inner.MeasureVolume(volume);
        public TokenLookup FindToken(ArchiveVolume volume, ArchiveItem item) => Inner.FindToken(volume, item);
        public DestinationIdentity Import(ArchiveVolume volume, ArchiveItem item)
        { counters.LegacyCalls++; SaveCounters(); return Inner.Import(volume, item); }
        public ArchiveSnapshot TargetSnapshot(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity, string directory, bool requireToken = true) => Profile(Inner.TargetSnapshot(volume, item, identity, directory, requireToken), directory);
        public List<DestinationIdentity> UnmarkedCandidates(ArchiveVolume volume, ArchiveItem item) => Inner.UnmarkedCandidates(volume, item);
        public void MarkCopy(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity) => Inner.MarkCopy(volume, item, identity);
        public DiskRoundtrip ReopenVolume(ArchiveVolume volume) { ReopenCalls++; return Inner.ReopenVolume(volume); }
        public SourceLocation? LocateSource(ArchiveItem item) => Inner.LocateSource(item);
        public SourceLocation MoveToDeletedItems(ArchiveItem item, string? expected) => Inner.MoveToDeletedItems(item, expected);
        public void DeleteFromDeletedItems(ArchiveItem item, SourceLocation location, string? expected) => Inner.DeleteFromDeletedItems(item, location, expected);
        public void Dispose() => Inner.Dispose();
    }
}
