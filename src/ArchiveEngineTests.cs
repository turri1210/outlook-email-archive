using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

// Actual SQLite/filesystem with a simulated mail provider. No Outlook/real mail.
internal static class ArchiveEngineTests
{
    internal static int Run()
    {
        ArchiveVolumePolicy.SelfTest(); DeliveryReportTests.Run(); RepairEvidenceTests.Run(); ArchiveValidation.SelfTest(); EmbeddedMessageEvidence.SelfTest(); DateFilterTests(); IdentityAliasTests();
        string directory = Path.Combine(Path.GetTempPath(), "arc-tests-" + Guid.NewGuid().ToString("N")[..16]);
        Directory.CreateDirectory(directory);
        LedgerTests(Path.Combine(directory, "ledger.sqlite"));
        foreach (string point in new[] { "after-export", "after-import-intent", "after-import-before-checkpoint" })
            ResumeCopy(Path.Combine(directory, point), point);
        ResumeRemoval(Path.Combine(directory, "delete-recovery"));
        AutomaticTrash(Path.Combine(directory,"auto-trash"));
        IsolatedOutlookFailure(Path.Combine(directory, "isolated-com"));
        BlockChangedSource(Path.Combine(directory, "source-change"));
        BlockDuplicatedToken(Path.Combine(directory, "token-collision"));
        BlockUnmarkedImport(Path.Combine(directory, "lost-token"));
        LegacyProfileReconciliation(Path.Combine(directory, "legacy-profile"));
        BlockAmbiguousUnmarkedImport(Path.Combine(directory, "ambiguous"));
        FreshImportPreservesUnrelatedUnmarked(Path.Combine(directory, "unrelated-unmarked"));
        ResumedIntentWithUnrelatedUnmarkedStaysBlocked(Path.Combine(directory, "resumed-unrelated-unmarked"));
        PreserveIdenticalMessages(Path.Combine(directory, "identical"));
        BlockMissingDurability(Path.Combine(directory, "durability"));
        ExplicitRetainedBackup(Path.Combine(directory, "explicit-backup"));
        ReadingDoesNotChangeArchivedContent(Path.Combine(directory, "read-status"));
        FolderStructurePersistence(Path.Combine(directory, "folder-structure"));
        IncrementalVerification(Path.Combine(directory, "incremental-verification"));
        EmbeddedLegacyRecovery(Path.Combine(directory, "embedded-legacy"));
        ExplicitRepairRecovery(Path.Combine(directory, "explicit-repair"));
        MissingSourceBackupProof(Path.Combine(directory, "missing-source-proof"));
        DeliveryReports(Path.Combine(directory,"ndr"));
        IgnoredMeetingPolicy(Path.Combine(directory,"ignored-meetings"));
        PersistentVerificationFailure(Path.Combine(directory, "persistent-verification"));
        ResumeTransientReopen(Path.Combine(directory, "reopen"));
        ExistingPstResume(Path.Combine(directory, "existing"));
        ExistingPstRotation(Path.Combine(directory, "existing-roll"));
        ChosenPstDestination(Path.Combine(directory,"chosen-pst"));
        ExistingPstConfigGuards(Path.Combine(directory, "pst-guards"));
        FrozenDateFilteredPlan(Path.Combine(directory, "date-plan"));
        FrozenScopeAndNewArrivals(Path.Combine(directory, "arrivals"));
        FrozenScopeMissingItem(Path.Combine(directory, "scope-missing"));
        OwnTrashMoveIsNotArrival(Path.Combine(directory, "own-trash"));
        SameChangeKeySkipsOnlyImportSnapshot(Path.Combine(directory, "key-same"));
        ChangedKeyAndContentBlocksImport(Path.Combine(directory, "key-content"));
        ChangedKeySameContentUsesFullSnapshot(Path.Combine(directory, "key-only"));
        MissingChangeKeyUsesFullSnapshot(Path.Combine(directory, "key-missing"));
        CatalogLease(Path.Combine(directory, "lease"));
        BlockChangedBackup(Path.Combine(directory, "backup-change"));
        Console.WriteLine("Motor: SQLite, deduplicação por identidade, mensagens iguais preservadas, rotação, retomada, durabilidade, bloqueio concorrente, origem/backup alterados e candidatos ambíguos conferidos.");
        Console.WriteLine("Evidências sintéticas preservadas em " + directory);
        return 0;
    }
    private static ArchiveConfig Config(string root) => new()
    {
        RootDirectory = root, SourceStoreId = "source", SourceFolderId = "folder", IncludeSubfolders = true,
        TargetVolumeBytes = 2 * ArchiveVolumePolicy.MiB, HardLimitBytes = 12 * ArchiveVolumePolicy.MiB, ReserveBytes = 128 * 1024
    };
    private static void IdentityAliasTests()
    {
        int comparisons = 0;
        bool CompareAlias(string left, string right)
        {
            comparisons++;
            return left == "short-term-alias" && right == "canonical-long-term-id";
        }
        Check(OutlookIdentity.AreSame("A1B2C3", "A1B2C3", (_, _) => throw new Exception("Comparação literal não deve chamar COM.")),
            "Identidades iguais não usaram a comparação literal sem callback.");
        Check(OutlookIdentity.AreSame("a1b2c3d4", "A1B2C3D4", (_, _) => throw new Exception("Capitalização hexadecimal não deve chamar COM.")),
            "Capitalização hexadecimal de uma identidade canônica alterou sua igualdade.");
        Check(OutlookIdentity.AreSame("short-term-alias", "canonical-long-term-id", CompareAlias) && comparisons == 1,
            "Alias diferente mas semanticamente igual foi rejeitado ou comparado repetidamente.");
        Check(!OutlookIdentity.AreSame("different-message-id", "canonical-long-term-id", CompareAlias) && comparisons == 2,
            "Identidades diferentes com callback negativo foram aceitas.");
        foreach (var pair in new (string? Left, string? Right)[]
        {
            (null, null), (null, "A1B2"), ("A1B2", null), ("", ""), ("", "A1B2"), ("A1B2", ""), (" ", "A1B2"), ("A1B2", " ")
        })
            Check(!OutlookIdentity.AreSame(pair.Left, pair.Right, (_, _) => throw new Exception("Identidade ausente não deve chamar COM.")),
                "Identidade ausente/vazia foi aceita ou comparada por callback.");
        bool propagated = false;
        try { OutlookIdentity.AreSame("unresolved-alias", "canonical-long-term-id", (_, _) => throw new IOException("Falha artificial de comparação MAPI.")); }
        catch (IOException) { propagated = true; }
        Check(propagated, "Falha na comparação semântica foi ocultada em vez de bloquear a operação.");
        Console.WriteLine("Identidades MAPI: igualdade literal/hexadecimal, aliases semânticos, divergências e falhas de comparação verificadas.");
    }
    private static void DateFilterTests()
    {
        const string saoPaulo = "E. South America Standard Time";
        const string eastern = "Eastern Standard Time";
        DateTime received = new(2025, 1, 1, 3, 0, 0, DateTimeKind.Utc);
        DateTime submitted = received.AddMinutes(-1);
        Check(ArchiveDateFilter.SelectDate(received, submitted) == (received, "received"), "Mensagem enviada a si próprio perdeu a data de recebimento.");
        Check(ArchiveDateFilter.SelectDate(null, submitted) == (submitted, "sent"), "Mensagem sem recebimento não usa envio.");
        Check(ArchiveDateFilter.SelectDate(new DateTime(4501, 1, 1), submitted) == (submitted, "sent"), "Data sentinela de recebimento não usa envio válido.");
        Check(ArchiveDateFilter.SelectDate(null, null).Utc is null, "Datas ausentes inventaram um instante.");
        static DateTime Utc(int year, int month, int day, int hour = 0, int minute = 0, int second = 0) =>
            new(year, month, day, hour, minute, second, DateTimeKind.Utc);
        static bool Includes(DateTime? utc, DateOnly? cutoff, string zone = "UTC") => ArchiveDateFilter.Includes(utc, cutoff, zone);
        static void RejectDate(DateTime? value)
        {
            try { Includes(value, new DateOnly(2026, 12, 31)); }
            catch (ArchiveReviewException) { return; }
            throw new Exception("Data ausente/sentinela não bloqueou o filtro ativo.");
        }
        Check(Includes(null, null), "Compatibilidade sem corte rejeitou mensagem sem data.");
        Check(Includes(Utc(4500, 1, 1), null), "Compatibilidade sem corte inspecionou data sentinela.");
        RejectDate(null); RejectDate(Utc(4500, 1, 1)); RejectDate(Utc(1600, 12, 31, 23, 59, 59));
        Check(Includes(Utc(1601, 1, 1), new DateOnly(1601, 1, 1)), "Limite inferior válido de data MAPI foi rejeitado.");

        var yearEnd = new DateOnly(2025, 12, 31);
        DateTime lastLocalMillisecond = Utc(2026, 1, 1, 2, 59, 59).AddMilliseconds(999);
        Check(Includes(lastLocalMillisecond, yearEnd, saoPaulo), "Último milissegundo de 31/12 local foi excluído.");
        Check(Includes(Utc(2026, 1, 1, 3).AddTicks(-1), yearEnd, saoPaulo), "Último tick do dia inclusivo foi excluído.");
        Check(!Includes(Utc(2026, 1, 1, 3), yearEnd, saoPaulo), "Primeiro instante de 1/1 local foi incluído.");
        Check(!Includes(lastLocalMillisecond, yearEnd, "UTC"), "Filtro ignorou o fuso persistido e usou a data de São Paulo em UTC.");
        Check(Includes(Utc(2025, 12, 31, 23, 59, 59), yearEnd, "UTC"), "Último segundo do ano em UTC foi excluído.");
        DateTime unspecified = DateTime.SpecifyKind(lastLocalMillisecond, DateTimeKind.Unspecified);
        Check(Includes(unspecified, yearEnd, saoPaulo), "Data MAPI sem Kind não foi interpretada como UTC.");
        Check(!Includes(DateTime.SpecifyKind(Utc(2026, 1, 1, 3), DateTimeKind.Unspecified), yearEnd, saoPaulo),
            "UTC sem Kind foi confundido com horário local na mudança de dia.");

        var leapDay = new DateOnly(2024, 2, 29);
        Check(Includes(Utc(2024, 2, 29, 23, 59, 59).AddMilliseconds(999), leapDay), "Último milissegundo do dia bissexto foi excluído.");
        Check(!Includes(Utc(2024, 3, 1), leapDay), "Primeiro instante de março passou pelo corte de 29/2.");

        // Explicit time zones make these tests independent of the computer's
        // current zone. DST transitions produce 23-hour and 25-hour local days.
        var springDay = new DateOnly(2024, 3, 10);
        Check(Includes(Utc(2024, 3, 10, 5), springDay, eastern), "Início do dia de 23 horas foi excluído.");
        Check(Includes(Utc(2024, 3, 11, 4).AddTicks(-1), springDay, eastern), "Fim do dia de 23 horas foi excluído.");
        Check(!Includes(Utc(2024, 3, 11, 4), springDay, eastern), "Filtro supôs 24 horas no dia de início do horário de verão.");
        var fallDay = new DateOnly(2024, 11, 3);
        Check(Includes(Utc(2024, 11, 3, 4), fallDay, eastern), "Início do dia de 25 horas foi excluído.");
        Check(Includes(Utc(2024, 11, 3, 5, 30), fallDay, eastern) && Includes(Utc(2024, 11, 3, 6, 30), fallDay, eastern),
            "Uma ocorrência do horário local repetido na mudança de fuso foi excluída.");
        Check(Includes(Utc(2024, 11, 4, 5).AddTicks(-1), fallDay, eastern), "Fim do dia de 25 horas foi excluído.");
        Check(!Includes(Utc(2024, 11, 4, 5), fallDay, eastern), "Filtro supôs 24 horas no dia de término do horário de verão.");
        Console.WriteLine("Filtro de data: casos de ano novo, fuso persistido, data bissexta, dias DST de 23/25 horas e datas ausentes/sentinela verificados.");
    }
    private static void AutomaticTrash(string directory)
    {
        foreach(string mode in new[]{"default","changed","durability","after-delete-intent","after-trash-move-before-checkpoint","move-failure","connection"})
        {
            string folder=Path.Combine(directory,mode),root=Path.Combine(folder,"catalog");Directory.CreateDirectory(folder);
            using var provider=new FakeArchiveBackend(Path.Combine(folder,"mail.json"));
            using var engine=new ArchiveEngine(root,provider);var config=Config(root);config.MoveValidatedToDeletedItems=mode!="default";
            string job=engine.Plan(config);engine.Copy(job);
            if(mode=="default")
            {
                engine.Verify(job,pendingOnly:true);
                Check(engine.Items(job).All(x=>x.State==ArchiveItemState.Verified)&&provider.Sources.All(x=>x.Folder=="folder"),"Plano antigo moveu originais sem autorização.");
                int reads=provider.SourceSnapshotCalls;engine.Verify(job,pendingOnly:true,moveValidatedToDeletedItems:true);
                Check(provider.SourceSnapshotCalls==reads+5,"Movimento repetiu a conferência integral já feita nesta passagem.");
            }
            else if(mode=="changed")
            {
                provider.Sources[0].Fields["Subject"]=FakeArchiveBackend.Hash("changed");provider.Save();
                engine.Verify(job,pendingOnly:true);
                Check(engine.Items(job).Count(x=>x.State==ArchiveItemState.ReviewRequired)==1 && provider.Sources[0].Folder=="folder" && provider.Sources.Count(x=>x.Folder=="trash")==4,"Divergência não protegeu o original ou bloqueou itens válidos.");
                Check(provider.DeleteCalls==0,"Modo automático excluiu definitivamente.");continue;
            }
            else if(mode=="durability")
            {
                provider.ExclusiveDiskRead=false;engine.Verify(job,pendingOnly:true);
                Check(provider.Sources.All(x=>x.Folder=="folder"),"PST sem prova durável permitiu movimento.");
                engine.Verify(job,pendingOnly:true,allowRetainedBackup:true);
            }
            else if(mode=="move-failure")
            {
                provider.FailMoveCount=1;engine.Verify(job,pendingOnly:true);
                Check(engine.Items(job).Count(x=>x.State==ArchiveItemState.DeleteStarted)==1,"Falha de movimento perdeu a intenção durável.");
                engine.Verify(job,pendingOnly:true);
            }
            else if(mode=="connection")
            {
                provider.FailMoveCount=100;engine.Verify(job,pendingOnly:true);
                Check(engine.Items(job).Count(x=>x.State==ArchiveItemState.DeleteStarted)==3 && provider.Sources.All(x=>x.Folder=="folder"),"Falhas consecutivas não interromperam a passagem automática.");
                provider.FailMoveCount=0;engine.Verify(job,pendingOnly:true);
            }
            else
            {
                engine.Fault=(phase,_)=>{if(phase==mode)throw new ArchiveInjectedCrashException(mode);};
                try{engine.Verify(job,pendingOnly:true);throw new Exception("Interrupção automática não acionada.");}catch(ArchiveInjectedCrashException){}
                Check(engine.Items(job).Any(x=>x.State==ArchiveItemState.DeleteStarted),"Interrupção perdeu o checkpoint do movimento.");
                engine.Fault=null;engine.Verify(job,pendingOnly:true);
            }
            Check(engine.Items(job).All(x=>x.State==ArchiveItemState.InDeletedItems)&&provider.Sources.All(x=>x.Folder=="trash"&&!x.Removed)&&provider.DeleteCalls==0,"Movimento automático não concluiu com segurança.");
            int before=provider.SourceSnapshotCalls,reopens=provider.ReopenCalls;engine.Verify(job,pendingOnly:true,moveValidatedToDeletedItems:true);
            Check(provider.SourceSnapshotCalls==before&&provider.ReopenCalls==reopens,"Retomada repetiu trabalho de mensagens já movidas.");
        }
        Console.WriteLine("Movimento automático: consentimento, conferência única, divergências, durabilidade, interrupções e retomada sem exclusão definitiva aprovados.");
    }
    private static void LedgerTests(string path)
    {
        using var ledger = new ArchiveLedger(path);
        var job = new ArchiveJob { Id = Guid.NewGuid().ToString("N"), RootDirectory = Path.GetDirectoryName(path)!, ConfigJson = "{}" }; ledger.InsertJob(job);
        var row = new ArchiveItem { JobId = job.Id, SourceStoreId = "a", SourceEntryId = "bc", SourceFolderId = "f" };
        Check(ledger.AddItem(row), "Registro inicial falhou.");
        Check(!ledger.AddItem(new ArchiveItem { JobId = job.Id, SourceStoreId = "a", SourceEntryId = "bc" }), "Duplicata de origem aceita.");
        Check(row.Key != ArchiveLedger.ComputeItemKey("ab", "c"), "Identidade sem enquadramento colidiu.");
        var stale = ledger.GetItem(row.Key)!; row.State = ArchiveItemState.Exported; ledger.UpdateItem(row);
        int count = ledger.GetEvents(job.Id).Count;
        MustFail(() => ledger.UpdateItem(stale));
        Check(ledger.GetEvents(job.Id).Count == count && ledger.GetItem(row.Key)!.State == ArchiveItemState.Exported, "Falha de versão não fez rollback.");
        var bad = ledger.GetItem(row.Key)!; bad.VolumeNumber = 999;
        MustFail(() => ledger.UpdateItem(bad));
        Check(ledger.GetItem(row.Key)!.VolumeNumber is null && ledger.GetEvents(job.Id).Count == count, "FK/evento não transacional.");
        var secondJob = new ArchiveJob { Id = Guid.NewGuid().ToString("N"), RootDirectory = Path.GetDirectoryName(path)!, ConfigJson = "{}" };
        ledger.InsertJob(secondJob);
        foreach (var identity in new[] { (Job: job.Id, Volume: 1), (Job: job.Id, Volume: 2), (Job: secondJob.Id, Volume: 1) })
            Check(ledger.AddVolume(new ArchiveVolume { JobId = identity.Job, Number = identity.Volume,
                PstPath = Path.Combine(Path.GetDirectoryName(path)!, identity.Job + "-" + identity.Volume + ".pst") }), "Volume de reserva não registrado.");
        row.VolumeNumber = 1; row.ReservedBytes = 0; ledger.UpdateItem(row);
        var twoBytes = new ArchiveItem { JobId = job.Id, SourceStoreId = "a", SourceEntryId = "reserved-two", VolumeNumber = 1, ReservedBytes = 2 };
        Check(ledger.AddItem(twoBytes), "Reserva de dois bytes não registrada.");
        Check(ledger.AddItem(new ArchiveItem { JobId = job.Id, SourceStoreId = "a", SourceEntryId = "reserved-other-volume", VolumeNumber = 2,
            ReservedBytes = 3L * int.MaxValue }), "Reserva de outro volume não registrada.");
        Check(ledger.AddItem(new ArchiveItem { JobId = secondJob.Id, SourceStoreId = "a", SourceEntryId = "reserved-other-job", VolumeNumber = 1,
            ReservedBytes = 73 }), "Reserva de outra execução não registrada.");
        Check(ledger.GetReservedBytes(job.Id, 1, row.Key) == 2 && ledger.GetReservedBytes(job.Id, 1, twoBytes.Key) == 0,
            "Agregação de reserva não excluiu a identidade correta ou ignorou o valor zero.");
        Check(ledger.GetReservedBytes(job.Id, 999, "not-an-item") == 0 && ledger.GetReservedBytes("unknown-job", 1, "not-an-item") == 0,
            "Agregação vazia não retornou zero.");
        foreach (string id in new[] { job.Id, secondJob.Id })
            foreach (int volume in new[] { 1, 2 })
                foreach (string excluded in new[] { row.Key, twoBytes.Key, "not-an-item" })
                    Check(ledger.GetReservedBytes(id, volume, excluded) == ledger.GetItems(id).Where(x => x.VolumeNumber == volume && x.Key != excluded).Sum(x => x.ReservedBytes),
                        "Soma SQL de reservas divergiu da soma anterior por execução, parte e identidade, inclusive em valores de 64 bits.");
        var trackedVolume = ledger.GetVolumes(job.Id).Single(x => x.Number == 1);
        int volumeEvents = ledger.GetEvents(job.Id).Count, unchangedVersion = trackedVolume.Version;
        ledger.UpdateVolume(trackedVolume);
        Check(trackedVolume.Version == unchangedVersion && ledger.GetEvents(job.Id).Count == volumeEvents,
            "Atualização sem mudança de volume incrementou versão ou criou evento redundante.");
        foreach (Action<ArchiveVolume> change in new Action<ArchiveVolume>[]
        {
            x => x.StoreId = "synthetic-target-store", x => x.State = ArchiveVolumeState.Ready,
            x => x.ObservedBytes = 271360, x => x.InitialBytes = 271360
        })
        {
            int previousVersion = trackedVolume.Version;
            change(trackedVolume); ledger.UpdateVolume(trackedVolume); volumeEvents++;
            Check(trackedVolume.Version == previousVersion + 1 && ledger.GetEvents(job.Id).Count == volumeEvents,
                "Mudança real de volume não preservou checkpoint e evento duráveis.");
            var staleVolume = JsonSerializer.Deserialize<ArchiveVolume>(JsonSerializer.Serialize(trackedVolume))!;
            staleVolume.Version = previousVersion;
            MustFail(() => ledger.UpdateVolume(staleVolume));
            Check(ledger.GetVolumes(job.Id).Single(x => x.Number == 1).Version == trackedVolume.Version && ledger.GetEvents(job.Id).Count == volumeEvents,
                "Versão obsoleta idêntica aos campos atuais passou pelo atalho de atualização sem mudança.");
            ledger.UpdateVolume(trackedVolume);
            Check(trackedVolume.Version == previousVersion + 1 && ledger.GetEvents(job.Id).Count == volumeEvents,
                "Repetição idêntica após mudança real gerou novo checkpoint de volume.");
        }
    }
    private static void ResumeCopy(string directory, string point)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"), state = Path.Combine(directory, "mail.json");
        using var provider = new FakeArchiveBackend(state);
        string job;
        using (var engine = new ArchiveEngine(root, provider))
        {
            job = engine.Plan(Config(root)); engine.Fault = (phase, _) => { if (phase == point) throw new ArchiveInjectedCrashException(point); };
            try { engine.Copy(job); throw new Exception("Interrupção não acionada."); } catch (ArchiveInjectedCrashException) { }
        }
        using (var engine = new ArchiveEngine(root, provider))
        {
            engine.Copy(job); engine.Verify(job);
            Check(provider.Copies.Count == 5 && provider.Copies.Select(x => x.Token).Distinct().Count() == 5, "Retomada duplicou cópias.");
            Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Cópias completas não conferidas.");
            Check(engine.Volumes(job).Count > 1, "Rotação não exercitada.");
            engine.Copy(job); Check(provider.Copies.Count == 5, "Reexecução duplicou cópias.");
            string second = engine.Plan(Config(root)); engine.Copy(second);
            Check(engine.Items(second).Count == 0 && provider.Copies.Count == 5, "Novo job duplicou origem já catalogada.");
        }
    }
    private static void ResumeRemoval(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"); using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        string job;
        using (var engine = new ArchiveEngine(root, provider))
        {
            job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
            engine.Fault = (phase, _) => { if (phase == "after-trash-move-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
            try { engine.Remove(job, true); throw new Exception("Interrupção de Move não acionada."); } catch (ArchiveInjectedCrashException) { }
        }
        using (var engine = new ArchiveEngine(root, provider))
        {
            engine.Fault = (phase, _) => { if (phase == "after-delete-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
            try { engine.Remove(job, true); throw new Exception("Interrupção de Delete não acionada."); } catch (ArchiveInjectedCrashException) { }
        }
        using (var engine = new ArchiveEngine(root, provider))
        {
            engine.Remove(job, true); Check(engine.Items(job).All(x => x.State == ArchiveItemState.Removed), "Remoção não reconciliou após crash.");
            Check(provider.Sources.All(x => x.Removed) && provider.Copies.Count == 5, "Remoção atingiu cópia ou não concluiu origem.");
            int deleted = provider.DeleteCalls; engine.Remove(job, true); Check(provider.DeleteCalls == deleted, "Remoção concluída foi repetida.");
        }
    }
    private static void BlockChangedSource(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"); using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
        provider.Sources[0].Fields["Subject"] = FakeArchiveBackend.Hash("alterado"); provider.Save();
        engine.Remove(job, true); Check(!provider.Sources[0].Removed && provider.DeleteCalls == 4, "Origem alterada foi removida.");
    }
    private static void BlockDuplicatedToken(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"); using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
        provider.Copies.Add(provider.Copies[0] with { EntryId = "duplicate-entry" }); provider.Save();
        engine.Remove(job, true); Check(!provider.Sources[0].Removed, "Token duplicado autorizou remoção.");
    }
    private static void BlockUnmarkedImport(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"); using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json")) { DropNextToken = true };
        string job;
        using (var engine = new ArchiveEngine(root, provider))
        {
            job = engine.Plan(Config(root)); engine.Fault = (phase, _) => { if (phase == "after-import-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
            try { engine.Copy(job); } catch (ArchiveInjectedCrashException) { }
        }
        using (var engine = new ArchiveEngine(root, provider))
        {
            engine.Copy(job); engine.Verify(job);
            Check(provider.Copies.Count(x => x.SourceId == "source-0") == 1, "Cópia sem token foi duplicada na retomada.");
            Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Candidato único e integral sem token não foi recuperado e validado.");
        }
    }
    private static void BlockAmbiguousUnmarkedImport(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json")) { DropNextToken = true };
        string job;
        using (var engine = new ArchiveEngine(root, provider))
        {
            job = engine.Plan(Config(root));
            engine.Fault = (phase, _) => { if (phase == "after-import-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
            try { engine.Copy(job); throw new Exception("Interrupção de importação não acionada."); } catch (ArchiveInjectedCrashException) { }
        }
        var first = provider.Copies.Single();
        provider.Copies.Add(first with { EntryId = "ambiguous-unmarked-entry" }); provider.Save();
        using var resumed = new ArchiveEngine(root, provider);
        resumed.Copy(job); resumed.Verify(job); resumed.Reconcile(job);
        var blocked = resumed.Items(job).Single(x => x.SourceEntryId == first.SourceId);
        Check(blocked.State == ArchiveItemState.ReviewRequired, "Dois candidatos integralmente iguais autorizaram adoção automática.");
        Check(provider.Copies.Count(x => x.SourceId == first.SourceId) == 2 && provider.Copies.Where(x => x.SourceId == first.SourceId).All(x => x.Token is null),
            "Ambiguidade marcou um candidato arbitrário ou criou uma terceira cópia.");
        resumed.Remove(job, true);
        Check(!provider.Sources.Single(x => x.Id == first.SourceId).Removed, "Origem ambígua foi removida.");
    }
    private static void FreshImportPreservesUnrelatedUnmarked(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider);
        var config = Config(root); config.TargetVolumeBytes = 10 * ArchiveVolumePolicy.MiB;
        string job = engine.Plan(config);
        engine.Fault = (phase, item) =>
        {
            if (phase != "after-import-intent" || provider.Copies.Count != 0) return;
            // Equal contents do not make a preexisting, unrelated identity the
            // copy of a source that this invocation has never imported.
            var source = provider.Sources.Single(x => x.Id == item.SourceEntryId);
            provider.Copies.Add(new(job, item.VolumeNumber!.Value, "unrelated-preexisting", null,
                "unrelated-entry", "target-folder", new(source.Fields), source.Size)); provider.Save();
        };
        engine.Copy(job); engine.Verify(job); engine.Copy(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified) && provider.Copies.Count == 6,
            "Uma cópia sem token não relacionada bloqueou importações novas ou foi adotada por conteúdo igual.");
        Check(provider.Copies.Single(x => x.EntryId == "unrelated-entry").Token is null &&
            provider.Sources.All(source => provider.Copies.Count(copy => copy.SourceId == source.Id) == 1),
            "Importação nova marcou a mensagem não relacionada, colapsou identidade ou duplicou uma origem.");
        Check(engine.Items(job).All(x => x.TargetEntryId != "unrelated-entry") && provider.DeleteCalls == 0,
            "Mensagem não relacionada foi reivindicada ou origens foram removidas na cópia.");
    }
    private static void ResumedIntentWithUnrelatedUnmarkedStaysBlocked(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        string job, interruptedSource = "";
        using (var engine = new ArchiveEngine(root, provider))
        {
            var config = Config(root); config.TargetVolumeBytes = 10 * ArchiveVolumePolicy.MiB;
            job = engine.Plan(config);
            engine.Fault = (phase, item) =>
            {
                if (phase != "after-import-intent") return;
                interruptedSource = item.SourceEntryId;
                var fields = new Dictionary<string, string>(provider.Sources.Single(x => x.Id == item.SourceEntryId).Fields);
                fields["Subject"] = FakeArchiveBackend.Hash("unrelated changed subject");
                provider.Copies.Add(new(job, item.VolumeNumber!.Value, "unrelated-preexisting", null,
                    "unrelated-entry", "target-folder", fields, 100)); provider.Save();
                throw new ArchiveInjectedCrashException(phase);
            };
            try { engine.Copy(job); throw new Exception("Interrupção da intenção não acionada."); }
            catch (ArchiveInjectedCrashException) { }
        }
        using var resumed = new ArchiveEngine(root, provider);
        resumed.Copy(job); resumed.Verify(job); resumed.Copy(job);
        Check(resumed.Items(job).Single(x => x.SourceEntryId == interruptedSource).State == ArchiveItemState.ReviewRequired &&
            !provider.Copies.Any(x => x.SourceId == interruptedSource),
            "Intenção retomada sem correspondência criou uma cópia ou foi liberada automaticamente da revisão.");
        Check(resumed.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 4 && provider.Copies.Count == 5 &&
            provider.Copies.Single(x => x.EntryId == "unrelated-entry").Token is null,
            "Uma intenção ambígua contaminou novas identidades ou alterou a cópia não relacionada.");
    }
    private static void PreserveIdenticalMessages(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        // Two real mailbox entries can have byte-for-byte equal message fields.
        // Their distinct source identities still require two destination items.
        provider.Sources[1].Fields = new(provider.Sources[0].Fields); provider.Save();
        using var engine = new ArchiveEngine(root, provider);
        var config = Config(root); config.TargetVolumeBytes = 10 * ArchiveVolumePolicy.MiB;
        string job = engine.Plan(config); engine.Copy(job); engine.Verify(job); engine.Copy(job);
        Check(provider.Copies.Count == 5 && engine.Items(job).Count == 5 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "Conteúdo idêntico colapsou identidades distintas ou a reexecução duplicou cópias.");
        var same = provider.Copies.Where(x => x.SourceId is "source-0" or "source-1").ToArray();
        Check(same.Length == 2 && same[0].EntryId != same[1].EntryId && same[0].Token != same[1].Token && same[0].Volume == same[1].Volume,
            "Mensagens iguais não mantiveram duas identidades independentes na mesma parte.");
        engine.Remove(job, true);
        Check(provider.Sources.All(x => x.Removed) && provider.Copies.Count == 5, "Remoção de origens iguais afetou cópias independentes.");
    }
    private static void BlockMissingDurability(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json")) { ExclusiveDiskRead = false };
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root));
        engine.Copy(job); engine.Verify(job); engine.Remove(job, true);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Copied) && provider.DeleteCalls == 0 && provider.Sources.All(x => x.Folder == "folder"),
            "Falta de durabilidade promoveu a validado ou alterou origens.");
        provider.ExclusiveDiskRead = true; engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Nova evidência de disco não permitiu verificar a cópia.");
        provider.ExclusiveDiskRead = false; engine.Remove(job, true);
        Check(provider.DeleteCalls == 0 && provider.Sources.All(x => !x.Removed && x.Folder == "folder"),
            "Validação antiga autorizou remoção sem nova evidência de durabilidade.");
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.ReviewRequired), "Perda de durabilidade na remoção não foi bloqueada para revisão.");
    }
    private static void IsolatedOutlookFailure(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
        provider.FailMoveCount = 1;
        engine.Remove(job, false);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.InDeletedItems) == 4 &&
            engine.Items(job).Count(x => x.State == ArchiveItemState.DeleteStarted && x.Error!.Contains("0x80004005")) == 1,
            "Uma falha COM isolada bloqueou a fila ou perdeu o checkpoint e diagnóstico.");
        engine.Remove(job, false);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.InDeletedItems && x.Error is null), "Falha COM não retomou com segurança.");
        // A separate provider/catalog is used below to exercise the circuit breaker.
        string secondRoot = Path.Combine(directory, "systemic");
        using var secondProvider = new FakeArchiveBackend(Path.Combine(directory, "second.json"));
        using var second = new ArchiveEngine(secondRoot, secondProvider);
        string secondJob = second.Plan(Config(secondRoot)); second.Copy(secondJob); second.Verify(secondJob);
        secondProvider.FailMoveCount = 5; second.Remove(secondJob, false);
        Check(secondProvider.FailMoveCount == 2 && second.Items(secondJob).Count(x => x.State == ArchiveItemState.Verified) == 2,
            "Falhas consecutivas do Outlook não interromperam a sessão.");
    }
    private static void EmbeddedLegacyRecovery(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json")) { EmitRelevantMetadata = true };
        foreach (var source in provider.Sources) source.Fields["Attachment-1-Type"] = FakeArchiveBackend.Hash("5");
        provider.Save();
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job);
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
            foreach (var row in ledger.GetItems(job))
            {
                row.State = ArchiveItemState.ReviewRequired;
                // Reproduce old deferred transfers whose target baseline was
                // never approved; do not replace an existing production baseline.
                row.TargetSnapshotHash = null; row.TargetFieldsPath = null; row.TargetRelevantHash = null;
                ledger.UpdateItem(row);
            }
        provider.EmitEmbeddedEvidence = true;
        foreach (var source in provider.Sources) source.Fields["Attachment-1-Hash"] = FakeArchiveBackend.Hash("source container " + source.Id);
        foreach (var copy in provider.Copies.ToList()) provider.Copies[provider.Copies.IndexOf(copy)] = copy with
            { Fields = new(copy.Fields) { ["Attachment-1-Hash"] = FakeArchiveBackend.Hash("target container " + copy.SourceId) } };
        provider.Save(); engine.Reconcile(job); engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Contêiner de e-mail anexado bloqueou validação da evidência íntegra.");
        var row0 = engine.Items(job).Single(x => x.SourceEntryId == "source-0");
        var current = provider.SourceSnapshot(row0, null, Path.Combine(directory, "marked-check"));
        var replay = provider.BackupSnapshot(row0, Path.Combine(directory, "backup-replay-check"));
        Check(ArchivedSourceEvidence.Matches(row0, current, replay), "Conferência final de marcação rejeitou replay íntegro.");
        engine.Fault = (point, _) => { if (point == "after-trash-move-before-checkpoint") throw new ArchiveInjectedCrashException(point); };
        try { engine.Remove(job, false); throw new Exception("Interrupção de movimentação não acionada."); } catch (ArchiveInjectedCrashException) { }
        engine.Fault = null; engine.Remove(job, false);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.InDeletedItems), "Retomada do e-mail anexado não concluiu movimentação.");
        foreach (var copy in provider.Copies.ToList()) provider.Copies[provider.Copies.IndexOf(copy)] = copy with
            { Fields = new(copy.Fields) { ["Attachment-1-Hash"] = FakeArchiveBackend.Hash("container after reopen " + copy.SourceId) } };
        provider.Save();
        engine.Fault = (point, _) => { if (point == "after-delete-before-checkpoint") throw new ArchiveInjectedCrashException(point); };
        try { engine.Remove(job, true); throw new Exception("Interrupção de exclusão não acionada."); } catch (ArchiveInjectedCrashException) { }
        engine.Fault = null; engine.Remove(job, true);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Removed) && provider.DeleteCalls == 5 && provider.Copies.Count == 5,
            "Replay do e-mail anexado perdeu o checkpoint ou repetiu cópias/exclusões.");
        Console.WriteLine("E-mails anexados: legado, verificação, marcação e duas retomadas sem duplicatas aprovados.");
    }
    private static void ExplicitRepairRecovery(string directory)
    {
        Directory.CreateDirectory(directory); string root=Path.Combine(directory,"catalog");
        using var provider=new FakeArchiveBackend(Path.Combine(directory,"mail.json"));
        using var engine=new ArchiveEngine(root,provider); string job=engine.Plan(Config(root)); engine.Copy(job);
        var selected=engine.Items(job)[0]; string originalMsg=selected.StageMsgPath!, originalSha=ArchiveFiles.Sha(originalMsg);
        var oldCopy=provider.Copies.Single(x=>x.SourceId==selected.SourceEntryId);
        provider.Copies[provider.Copies.IndexOf(oldCopy)]=oldCopy with { Fields=new(oldCopy.Fields) { ["Subject"]=FakeArchiveBackend.Hash("bad old copy") } }; provider.Save();
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
        { var row=ledger.GetItems(job).Single(x=>x.Key==selected.Key); row.State=ArchiveItemState.ReviewRequired; ledger.UpdateItem(row); }
        engine.Fault=(point,_)=>{ if(point=="after-repair-copy-before-checkpoint") throw new ArchiveInjectedCrashException(point); };
        try { engine.ResolvePending(job,new HashSet<string>{selected.Key},true); throw new Exception("Crash do reparo não acionado."); } catch(ArchiveInjectedCrashException) { }
        Check(provider.Copies.Count==6,"Reparo não preservou a cópia anterior.");
        engine.Fault=null; engine.ResolvePending(job,new HashSet<string>{selected.Key},true); engine.Verify(job,pendingOnly:true);
        var repaired=engine.Items(job).Single(x=>x.Key==selected.Key);
        Check(repaired.State==ArchiveItemState.Verified && repaired.RepairToken is not null && repaired.RepairToken!=repaired.Token && provider.Copies.Count==6,
            "Retomada do reparo duplicou mensagem ou não validou a cópia nova.");
        Check(ArchiveFiles.Sha(originalMsg)==originalSha && provider.Copies.Any(x=>x.EntryId==oldCopy.EntryId) && File.Exists(repaired.RepairManifestPath),
            "Reparo alterou backup ou apagou cópia anterior.");
        engine.Remove(job,false); Check(engine.Items(job).All(x=>x.State==ArchiveItemState.InDeletedItems),"Remoção não usou o destino reparado.");
        engine.ResolvePending(job,new HashSet<string>{selected.Key},true); Check(provider.Copies.Count==6,"Reparo concluído foi repetido.");
    }
    private static void IgnoredMeetingPolicy(string directory)
    {
        Directory.CreateDirectory(directory);string root=Path.Combine(directory,"catalog");
        using var provider=new FakeArchiveBackend(Path.Combine(directory,"mail.json"));
        string[] classes=["IPM.Schedule.Meeting.Request","IPM.Schedule.Meeting.Resp.Pos","IPM.Schedule.Meeting.Canceled"];
        using var engine=new ArchiveEngine(root,provider);string job=engine.Plan(Config(root));
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
            for(int i=0;i<classes.Length;i++)
            {
                provider.Sources[i].Fields["MessageClass"]=FakeArchiveBackend.Hash(classes[i]);
                var row=ledger.GetItems(job).Single(x=>x.SourceEntryId==provider.Sources[i].Id);row.State=ArchiveItemState.SkippedUnsupported;ledger.UpdateItem(row);
            }
        provider.Save();engine.Copy(job);engine.Verify(job);engine.Remove(job,false);engine.Remove(job,true);
        Check(engine.Items(job).Count(x=>x.State==ArchiveItemState.SkippedByPolicy)==3 && engine.Items(job).Count(x=>x.State==ArchiveItemState.Removed)==2,"Reuniões desconsideradas entraram na cópia ou exclusão.");
        Check(provider.Sources.Take(3).All(x=>!x.Removed && x.Folder=="folder"),"Original de reunião foi modificado.");
    }
    private static void DeliveryReports(string directory)
    {
        Directory.CreateDirectory(directory);string root=Path.Combine(directory,"catalog");
        using var provider=new FakeArchiveBackend(Path.Combine(directory,"mail.json"));
        foreach(var source in provider.Sources)
        {
            source.Fields["MessageClass"]=FakeArchiveBackend.Hash(DeliveryReportEvidence.MessageClass);
            foreach(string key in DeliveryReportEvidence.RequiredFields)source.Fields[key]=FakeArchiveBackend.Hash(key=="Report-Version"?"ndr-v1":key);
        }
        provider.Save();using var engine=new ArchiveEngine(root,provider);string job=engine.Plan(Config(root));
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
        { foreach(var row in ledger.GetItems(job)){row.State=ArchiveItemState.SkippedUnsupported;ledger.UpdateItem(row);} }
        engine.Copy(job);engine.Verify(job);
        Check(engine.Items(job).All(x=>x.State==ArchiveItemState.Verified),"NDR não validou no fluxo completo.");
        provider.Sources[0].Fields["Report-Recipients"]=FakeArchiveBackend.Hash("diagnóstico alterado");provider.Save();
        engine.Remove(job,false);
        Check(engine.Items(job).Count(x=>x.State==ArchiveItemState.InDeletedItems)==4 && provider.DeleteCalls==0,"NDR alterado foi removido ou NDR íntegro bloqueou.");
        engine.Remove(job,true);
        Check(engine.Items(job).Count(x=>x.State==ArchiveItemState.Removed)==4 && provider.DeleteCalls==4,"Exclusão definitiva NDR não respeitou o conjunto validado.");
    }
    private static void MissingSourceBackupProof(string directory)
    {
        Directory.CreateDirectory(directory); string root=Path.Combine(directory,"catalog");
        using var provider=new FakeArchiveBackend(Path.Combine(directory,"mail.json"));
        using var engine=new ArchiveEngine(root,provider); string job=engine.Plan(Config(root)); engine.Copy(job);
        var rows=engine.Items(job); var keys=rows.Select(x=>x.Key).ToHashSet();
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
            foreach(var row in ledger.GetItems(job)) { row.State=ArchiveItemState.ReviewRequired; ledger.UpdateItem(row); }
        var missingCopy=provider.Copies[0]; provider.Copies.RemoveAt(0);
        using(var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))
        {
            var missing=ledger.GetItems(job).Single(x=>x.SourceEntryId==missingCopy.SourceId);
            missing.TargetEntryId=null; missing.TargetFolderId=null; missing.TargetStoreId=null;
            missing.TargetSnapshotHash=null; missing.TargetFieldsPath=null; missing.TargetRelevantHash=null; ledger.UpdateItem(missing);
        }
        foreach(var source in provider.Sources) source.Removed=true; provider.Save();
        engine.ResolvePending(job,keys,true);
        Check(engine.Items(job).All(x=>x.State==ArchiveItemState.BackupVerifiedSourceUnmanaged) && provider.DeleteCalls==0,
            "Origem ausente não teve backup confirmado ou foi declarada excluída indevidamente.");
        engine.Remove(job,true); Check(provider.DeleteCalls==0,"Estado de origem desconhecida entrou na exclusão.");
    }
    private static void LegacyProfileReconciliation(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job);
        foreach (var copy in provider.Copies.ToList())
            provider.Copies[provider.Copies.IndexOf(copy)] = copy with { Token = null };
        provider.Save(); provider.EmitRelevantMetadata = true;
        provider.Sources[0].Fields["Subject"] = FakeArchiveBackend.Hash("alterado"); provider.Save();
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
            foreach (var row in ledger.GetItems(job)) { row.State = ArchiveItemState.ReviewRequired; ledger.UpdateItem(row); }
        string selectedKey = engine.Items(job).Single(x => x.SourceEntryId == "source-1").Key;
        engine.Reconcile(job, new HashSet<string> { selectedKey });
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Copied) == 1,
            "Reconciliação selecionada processou itens fora do escopo.");
        engine.Reconcile(job);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Copied) == 4 &&
            engine.Items(job).Single(x => x.SourceEntryId == "source-0").State == ArchiveItemState.ReviewRequired && provider.Copies.Count == 5,
            "Perfil legado impediu recuperação ou aprovou original alterado.");
    }
    private static void PersistentVerificationFailure(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job);
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
        {
            var row = ledger.GetItems(job)[0];
            row.Error = "verification-pending:InvalidOperationException:Anexo especial não suportado no piloto; revisão necessária.";
            ledger.UpdateItem(row);
        }
        int reads = provider.SourceSnapshotCalls;
        engine.Verify(job, pendingOnly: true);
        Check(provider.SourceSnapshotCalls == reads + 4 && engine.Items(job).Count(x => x.State == ArchiveItemState.ReviewRequired) == 1,
            "Falha persistente foi repetida ou autorizou remoção.");
        reads = provider.SourceSnapshotCalls;
        engine.Verify(job, pendingOnly: true);
        Check(provider.SourceSnapshotCalls == reads, "Falha persistente retornou à fila automática.");
    }
    private static void IncrementalVerification(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job);
        engine.Verify(job, maxItems: 2, pendingOnly: true);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 2, "Lote inicial de validação não respeitou o limite.");
        int reads = provider.SourceSnapshotCalls;
        provider.Sources[0].Fields["Subject"] = FakeArchiveBackend.Hash("alterado após validação"); provider.Save();
        engine.Verify(job, pendingOnly: true);
        Check(provider.SourceSnapshotCalls == reads + 3 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Validação incremental releu itens já aprovados ou ignorou pendentes.");
        reads = provider.SourceSnapshotCalls; int reopen = provider.ReopenCalls;
        engine.Verify(job, pendingOnly: true);
        Check(provider.SourceSnapshotCalls == reads && provider.ReopenCalls == reopen, "Execução sem pendentes releu mensagens ou reabriu o PST.");
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
        {
            var oldWarning = ledger.GetItems(job).Single(x => x.SourceEntryId == "source-1");
            oldWarning.Error = "volume-reopen-pending:aviso antigo"; ledger.UpdateItem(oldWarning);
        }
        engine.Remove(job, false);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.InDeletedItems && x.Error is null) == 4 &&
            engine.Items(job).Count(x => x.State == ArchiveItemState.ReviewRequired) == 1 && provider.SourceSnapshotCalls > reads &&
            provider.Sources[0].Folder == "folder", "Remoção não revalidou o original alterado ou manteve avisos antigos nos itens aprovados.");
        reads = provider.SourceSnapshotCalls; reopen = provider.ReopenCalls;
        engine.Remove(job, false);
        Check(provider.SourceSnapshotCalls == reads && provider.ReopenCalls == reopen,
            "Retomada releu mensagens já movidas para Itens Excluídos.");
        engine.Remove(job, true);
        Check(provider.SourceSnapshotCalls == reads + 4 && engine.Items(job).Count(x => x.State == ArchiveItemState.Removed) == 4,
            "Remoção final não fez exatamente uma conferência por item já na lixeira.");
        Console.WriteLine("Validação incremental: 3 pendentes lidos, 2 aprovados pulados, repetição sem leituras e alteração posterior bloqueada na remoção.");
    }
    private static void FolderStructurePersistence(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider);
        var config = Config(root); config.PreserveFolderStructure = true;
        config.FolderMappings["[\"Caixa de Entrada\",\"AC BR\"]"] = new(6, "Caixa de Entrada", ["AC BR"]);
        string job = engine.Plan(config); engine.Copy(job);
        using var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite"), true);
        var volume = ledger.GetVolumes(job).First();
        Check(volume.PreserveFolderStructure && volume.FolderMappings.Single().Value.Subfolders.SequenceEqual(new[] { "AC BR" }), "Mapeamento não persistiu no volume.");
        Check(!JsonSerializer.Deserialize<ArchiveConfig>("{}")!.PreserveFolderStructure && !JsonSerializer.Deserialize<ArchiveVolume>("{}")!.PreserveFolderStructure, "Execuções antigas trocaram de organização.");
        Check(FolderStructureRules.StandardNames(6, "Caixa de Entrada").Contains("Inbox") && FolderStructureRules.StandardNames(5, "Itens Enviados").Contains("Sent Items"), "Equivalência de idiomas incompleta.");
        string wrapped = Convert.ToHexString(new byte[] { 9, 8, 7 }.Concat(System.Text.Encoding.Unicode.GetBytes("C:\\Archive\\real.pst\0")).ToArray());
        Check(FolderStructureRules.StoreIdContainsPath(wrapped, "c:\\archive\\REAL.pst"), "Identificador do PST novo não reconhecido.");
        Check(!FolderStructureRules.StoreIdContainsPath(wrapped, "C:\\Archive\\real") && !FolderStructureRules.StoreIdContainsPath(wrapped, "C:\\Other\\real.pst"), "Outro arquivo foi aceito pelo identificador.");
    }
    private static void ReadingDoesNotChangeArchivedContent(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        foreach (var source in provider.Sources) source.Fields["UnRead"] = FakeArchiveBackend.Hash("True");
        provider.Save(); using var engine = new ArchiveEngine(root, provider);
        string job = engine.Plan(Config(root)); engine.Copy(job);
        foreach (var source in provider.Sources) source.Fields["UnRead"] = FakeArchiveBackend.Hash("False");
        provider.Save();
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite")))
            foreach (var row in ledger.GetItems(job)) { row.State = ArchiveItemState.ReviewRequired; row.Error = "read-status"; ledger.UpdateItem(row); }
        engine.Reconcile(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Ler o original bloqueou a reconciliação.");
        engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified), "Ler o original bloqueou a validação.");
        var sample = engine.Items(job).First();
        var current = provider.SourceSnapshot(sample, null, Path.Combine(directory, "preparation-current"));
        Check(ArchivedSourceEvidence.Matches(sample, current), "Preparação da remoção rejeitou apenas UnRead.");
        foreach (string field in new[] { "Subject", "Attachment-1-Hash", "NativeBody-MSGBodyStream:__substg1.0_1000001F" })
        {
            var changed = new Dictionary<string, string>(current.Fields) { [field] = FakeArchiveBackend.Hash("alterado") };
            Check(!ArchivedSourceEvidence.Matches(sample, current with { Fields = changed }), "Preparação da remoção ignorou alteração de " + field);
        }
        engine.Remove(job, false);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.InDeletedItems), "Leitura bloqueou remoção sintética validada.");
        engine.Fault = (phase, _) => { if (phase == "after-delete-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
        try { engine.Remove(job, true); throw new Exception("Interrupção sintética não acionada."); } catch (ArchiveInjectedCrashException) { }
        engine.Fault = null; engine.Remove(job, true);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Removed), "Leitura bloqueou retomada após remoção já executada.");
    }
    private static void ExplicitRetainedBackup(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json")) { ExclusiveDiskRead = false };
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root));
        engine.Copy(job); engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Copied), "Prova alternativa foi usada sem autorização.");
        provider.Sources[0].Fields["Subject"] = FakeArchiveBackend.Hash("alterado"); provider.Save();
        engine.Verify(job, allowRetainedBackup: true);
        Check(engine.Items(job).Count(x => x.State == ArchiveItemState.Verified) == 4 && engine.Items(job).Count(x => x.State == ArchiveItemState.ReviewRequired) == 1, "Autorização explícita não separou itens íntegros e divergentes.");
        using (var ledger = new ArchiveLedger(Path.Combine(root, "archive.sqlite"), true))
            Check(!JsonSerializer.Deserialize<ArchivePlanEnvelope>(ledger.GetJob(job).ConfigJson)!.Config.AllowRetainedMsgBackupForRemoval, "Autorização alterou o plano congelado.");
        engine.Remove(job, true, allowRetainedBackup: true);
        Check(provider.Sources.Count(x => x.Removed) == 4 && !provider.Sources[0].Removed && provider.Sources[0].Folder == "folder", "Remoção parcial afetou a origem divergente ou bloqueou as demais.");
    }
    private static void CatalogLease(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        string job;
        using (var first = new ArchiveEngine(root, provider))
        {
            job = first.Plan(Config(root)); bool blocked = false;
            try { using var second = new ArchiveEngine(root, provider); }
            catch (IOException) { blocked = true; }
            Check(blocked, "Duas instâncias obtiveram escrita simultânea no mesmo catálogo.");
            using var reader = new ArchiveLedger(Path.Combine(root, "archive.sqlite"), true);
            Check(reader.GetItems(job).Count == 5, "Consulta de status foi bloqueada pelo executor ativo.");
        }
        using var resumed = new ArchiveEngine(root, provider);
        Check(resumed.Items(job).Count == 5, "Liberação do bloqueio perdeu o catálogo ou impediu nova instância.");
    }
    private static void ResumeTransientReopen(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root));
        engine.Verify(job); engine.Remove(job, true);
        Check(provider.ReopenCalls == 0, "Plano sem itens elegíveis provocou reabertura desnecessária.");
        engine.Copy(job); provider.ThrowOnReopen = true; engine.Verify(job);
        Check(provider.ReopenCalls > 0 && engine.Items(job).All(x => x.State == ArchiveItemState.Copied && x.Error?.StartsWith("volume-reopen-pending:", StringComparison.Ordinal) == true),
            "Falha transitória de reabertura não preservou Copied com erro visível.");
        Check(engine.Volumes(job).All(x => x.State == ArchiveVolumeState.Ready), "Falha transitória congelou a parte para revisão.");
        provider.ThrowOnReopen = false; engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified && x.Error is null), "Nova tentativa não verificou e limpou a falha transitória.");
        provider.ThrowOnReopen = true; engine.Verify(job);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Verified && x.Error?.StartsWith("volume-reopen-pending:", StringComparison.Ordinal) == true),
            "Reverificação transitória rebaixou estado ou ocultou o erro.");
        engine.Remove(job, true);
        Check(provider.DeleteCalls == 0 && provider.Sources.All(x => x.Folder == "folder" && !x.Removed) &&
            engine.Items(job).All(x => x.State == ArchiveItemState.Verified && x.Error?.StartsWith("volume-reopen-pending:", StringComparison.Ordinal) == true),
            "Remoção com falha transitória mudou origem/estado ou não registrou o erro.");
        provider.ThrowOnReopen = false; engine.Remove(job, true);
        Check(engine.Items(job).All(x => x.State == ArchiveItemState.Removed && x.Error is null), "Retomada da remoção não concluiu ou reteve erro transitório antigo.");
        int reopenCalls = provider.ReopenCalls, deleteCalls = provider.DeleteCalls;
        provider.ThrowOnReopen = true; engine.Verify(job); engine.Remove(job, true);
        Check(provider.ReopenCalls == reopenCalls && provider.DeleteCalls == deleteCalls,
            "Reexecução concluída reabriu partes ou repetiu remoções.");
    }
    private static byte[] WriteExistingPst(string path, int length = 271360, ushort version = 23)
    {
        // Artificial header and sentinel payload only. This is never opened in
        // Outlook and is intentionally not a complete PST implementation.
        var bytes = new byte[length];
        new byte[] { 0x21, 0x42, 0x44, 0x4E }.CopyTo(bytes, 0);
        bytes[8] = 0x53; bytes[9] = 0x4D;
        System.Buffers.Binary.BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(10, 2), version);
        Encoding.UTF8.GetBytes("Conteúdo manual preexistente — não sobrescrever.").CopyTo(bytes, 128);
        File.WriteAllBytes(path, bytes); return bytes;
    }
    private static void ExistingPstResume(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"), pst = Path.Combine(directory, "existing.pst");
        byte[] original = WriteExistingPst(pst);
        string alternate = Path.Combine(directory, "alternate.pst"); WriteExistingPst(alternate);
        string alternateHash = ArchiveFiles.Sha(alternate);
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        // A manually archived identical message belongs to neither this job nor
        // its deduplication scope, so it must not substitute for a new copy.
        provider.Copies.Add(new("manual", 1, "source-0", null, "manual-existing-entry", "manual-folder", new(provider.Sources[0].Fields), 0)); provider.Save();
        var config = Config(root); config.ExistingPstPath = pst; config.TargetVolumeBytes = 10 * ArchiveVolumePolicy.MiB;
        string job;
        using (var engine = new ArchiveEngine(root, provider))
        {
            job = engine.Plan(config);
            // Mutating the caller's configuration must not redirect the frozen
            // job to another PST after it has been planned.
            config.ExistingPstPath = alternate;
            engine.Fault = (phase, _) => { if (phase == "after-import-before-checkpoint") throw new ArchiveInjectedCrashException(phase); };
            try { engine.Copy(job); throw new Exception("Interrupção no PST existente não acionada."); } catch (ArchiveInjectedCrashException) { }
        }
        using var resumed = new ArchiveEngine(root, provider); resumed.Copy(job); resumed.Verify(job); resumed.Copy(job);
        var first = resumed.Volumes(job).Single(x => x.Number == 1);
        Check(first.IsExisting && first.SourceStoreId == "source" && first.PstPath == pst && resumed.Volumes(job).Count == 1,
            "Destino existente perdeu sinalização, identidade de origem ou caminho congelado.");
        Check(File.ReadAllBytes(pst).AsSpan(0, original.Length).SequenceEqual(original) && new FileInfo(pst).Length > original.Length,
            "Importação/retomada truncou ou sobrescreveu conteúdo anterior do PST.");
        Check(ArchiveFiles.Sha(alternate) == alternateHash, "Mutação de configuração redirecionou a execução para outro PST.");
        Check(provider.Copies.Count(x => x.JobId == job) == 5 && provider.Copies.Count == 6 && resumed.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "Mensagem manual foi adotada como deduplicação ou retomada duplicou as cinco mensagens novas.");
        Check(provider.Copies.Single(x => x.EntryId == "manual-existing-entry").Token is null, "Mensagem manual fora do escopo foi marcada pela automação.");
    }
    private static void ChosenPstDestination(string directory)
    {
        Directory.CreateDirectory(directory);string output=Path.Combine(directory,"dest");Directory.CreateDirectory(output);
        string selected=Path.Combine(output,"backup.pst");
        using var provider=new FakeArchiveBackend(Path.Combine(directory,"mail.json"));
        var config=Config(Path.Combine(directory,"catalog"));config.DestinationPstPath=selected;
        using var engine=new ArchiveEngine(config.RootDirectory,provider);string id=engine.Plan(config);
        config.DestinationPstPath=Path.Combine(directory,"wrong.pst");engine.Copy(id);engine.Verify(id);
        var volumes=engine.Volumes(id);
        Check(volumes.Count>1 && volumes[0].PstPath==selected && volumes.All(x=>Path.GetDirectoryName(x.PstPath)==output),"PST e rotação não respeitam a pasta escolhida.");
        Check(volumes.All(x=>!x.IsExisting) && engine.Items(id).All(x=>x.State==ArchiveItemState.Verified),"Novo PST não foi copiado e validado.");
        string hash=ArchiveFiles.Sha(selected);
        using var collisionProvider=new FakeArchiveBackend(Path.Combine(directory,"other-mail.json"));
        var collision=Config(Path.Combine(directory,"other"));collision.DestinationPstPath=selected;
        using var blocked=new ArchiveEngine(collision.RootDirectory,collisionProvider);string other=blocked.Plan(collision);blocked.Copy(other);
        Check(ArchiveFiles.Sha(selected)==hash && blocked.Items(other).All(x=>x.State==ArchiveItemState.ReviewRequired),"Colisão de nome sobrescreveu um PST existente.");
    }
    private static void ExistingPstRotation(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"), pst = Path.Combine(directory, "full.pst");
        WriteExistingPst(pst, (int)(3 * ArchiveVolumePolicy.MiB)); string hash = ArchiveFiles.Sha(pst);
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        var config = Config(root); config.ExistingPstPath = pst;
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(config); engine.Copy(job); engine.Verify(job);
        Check(engine.Volumes(job).Single(x => x.Number == 1).IsExisting && engine.Volumes(job).Count > 1,
            "PST existente acima da meta não foi registrado como primeira parte com rotação.");
        Check(provider.Copies.Count == 5 && provider.Copies.All(x => x.Volume >= 2) && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "PST já acima da meta recebeu novas mensagens ou rotação não concluiu.");
        Check(ArchiveFiles.Sha(pst) == hash, "Rotação alterou bytes do PST preexistente que deveria ficar sem importações.");
    }
    private static void ExistingPstConfigGuards(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        var config = Config(root); config.ExistingPstPath = Path.Combine(directory, "missing.pst");
        MustFail(config.Validate);
        Check(!File.Exists(config.ExistingPstPath), "Validação criou silenciosamente um PST ausente.");
        string ansi = Path.Combine(directory, "ansi.pst"); WriteExistingPst(ansi, version: 14); config.ExistingPstPath = ansi; MustFail(config.Validate);
        string invalid = Path.Combine(directory, "invalid.pst"); File.WriteAllBytes(invalid, new byte[271360]); config.ExistingPstPath = invalid; MustFail(config.Validate);
        string valid = Path.Combine(directory, "valid.pst"); WriteExistingPst(valid); config.ExistingPstPath = valid; config.Validate();
        string sync = Path.Combine(directory, "OneDrive"); Directory.CreateDirectory(sync);
        config.RootDirectory = Path.Combine(sync, "catalog"); MustFail(config.Validate);
        config.RootDirectory = root; string syncPst = Path.Combine(sync, "synced.pst"); WriteExistingPst(syncPst);
        config.ExistingPstPath = syncPst; MustFail(config.Validate);
    }
    private static void FrozenDateFilteredPlan(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        provider.Sources[0].MessageDateUtc = new DateTime(2025, 12, 31, 23, 59, 59, DateTimeKind.Utc);
        provider.Sources[1].MessageDateUtc = new DateTime(2026, 1, 1, 2, 59, 59, DateTimeKind.Utc).AddMilliseconds(999);
        provider.Sources[2].MessageDateUtc = new DateTime(2024, 2, 29, 12, 0, 0, DateTimeKind.Utc);
        provider.Sources[3].MessageDateUtc = new DateTime(2026, 1, 1, 3, 0, 0, DateTimeKind.Utc);
        provider.Sources[4].MessageDateUtc = new DateTime(2026, 1, 2, 0, 0, 0, DateTimeKind.Utc);
        provider.Save();
        var config = Config(root); config.CutoffDateInclusive = new DateOnly(2025, 12, 31);
        config.CutoffTimeZoneId = "E. South America Standard Time";
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(config);
        var expectedIds = new HashSet<string>(StringComparer.Ordinal) { "source-0", "source-1", "source-2" };
        Check(engine.Items(job).Count == 3 && expectedIds.SetEquals(engine.Items(job).Select(x => x.SourceEntryId)),
            "Plano por data não incluiu exatamente as três mensagens anteriores ao corte local inclusivo.");

        // Changes to the caller's config and messages arriving after planning
        // must never widen the immutable source plan during Copy or its replay.
        config.CutoffDateInclusive = new DateOnly(2026, 1, 31); config.CutoffTimeZoneId = "UTC";
        provider.Sources.Add(new()
        {
            Id = "arrived-after-plan", RecordKey = "record-arrived-after-plan",
            MessageDateUtc = new DateTime(2023, 1, 1, 0, 0, 0, DateTimeKind.Utc), Fields = new(provider.Sources[0].Fields)
        });
        provider.Sources.Add(new()
        {
            Id = "arrived-too-new", RecordKey = "record-arrived-too-new",
            MessageDateUtc = new DateTime(2026, 2, 1, 0, 0, 0, DateTimeKind.Utc), Fields = new(provider.Sources[0].Fields)
        });
        provider.Save();
        var arrivals = engine.CheckNew(job);
        Check(arrivals.Error is null && arrivals.NewItemsCount == 1,
            "Conferência de novas mensagens ignorou o corte congelado ou confundiu itens antigos excluídos com novas chegadas.");
        engine.Copy(job); engine.Verify(job); engine.Copy(job);
        Check(provider.Copies.Count == 3 && expectedIds.SetEquals(provider.Copies.Select(x => x.SourceId)),
            "Cópia/reexecução ampliou o escopo congelado após alteração de data/fuso ou chegada de nova mensagem.");
        Check(engine.Items(job).Count == 3 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "As três mensagens do plano por data não permaneceram únicas e verificadas.");
        Check(provider.Copies.All(x => x.SourceId is not ("source-3" or "source-4" or "arrived-after-plan" or "arrived-too-new")),
            "Mensagem posterior ao corte ou ausente no plano foi copiada.");
        Check(provider.DeleteCalls == 0 && provider.Sources.All(x => !x.Removed && x.Folder == "folder"),
            "Planejamento e cópia por data alteraram ou removeram originais.");
    }
    private static void FrozenScopeAndNewArrivals(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        var originalIds = provider.Sources.Select(x => x.Id).ToHashSet(StringComparer.Ordinal);
        provider.AfterCapture = () =>
        {
            foreach (string id in new[] { "new-first", "new-second" })
                provider.Sources.Add(new() { Id = id, RecordKey = "record-" + id, Fields = new(provider.Sources[0].Fields) });
            provider.Sources.Reverse(); provider.Save();
        };
        using var engine = new ArchiveEngine(root, provider); string original = engine.Plan(Config(root));
        Check(engine.Items(original).Count == 5 && originalIds.SetEquals(engine.Items(original).Select(x => x.SourceEntryId)),
            "Chegadas e reordenação após captura de IDs alteraram o plano inicial de cinco mensagens.");
        engine.Copy(original); engine.Verify(original);
        Check(provider.Copies.Count == 5 && originalIds.SetEquals(provider.Copies.Select(x => x.SourceId)),
            "Cópia do plano congelado incluiu mensagens novas ou perdeu mensagens originais após reordenação.");
        var arrival = engine.CheckNew(original);
        Check(arrival.Error is null && arrival.NewItemsCount == 2 && arrival.CheckedUtc != default,
            "Conferência não identificou exatamente as duas chegadas posteriores ao escopo inicial.");
        string next = engine.PlanNew(original);
        Check(next != original && engine.Items(next).Count == 2 && engine.Items(next).All(x => x.SourceEntryId is "new-first" or "new-second"),
            "Plano complementar incluiu mensagens do catálogo anterior ou perdeu chegadas.");
        var registered = engine.CheckNew(original);
        Check(registered.Error is null && registered.NewItemsCount == 0, "Conferência reapresentou chegadas já registradas em outro plano.");
        var fresh = engine.CheckNew(next);
        Check(fresh.Error is null && fresh.NewItemsCount == 0, "Escopo complementar não capturou a fotografia atual completa de IDs.");
        engine.Copy(next); engine.Verify(next); engine.Copy(next); engine.Copy(original);
        Check(provider.Copies.Count == 7 && provider.Copies.Select(x => x.SourceId).Distinct(StringComparer.Ordinal).Count() == 7 &&
            engine.Items(next).All(x => x.State == ArchiveItemState.Verified),
            "Retomada do plano complementar duplicou mensagens ou não concluiu sete destinos independentes.");
        Check(provider.DeleteCalls == 0 && provider.Sources.All(x => !x.Removed && x.Folder == "folder"), "Conferência de chegadas alterou originais.");
    }
    private static void FrozenScopeMissingItem(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        var config = Config(root); var scope = provider.CaptureScope(config);
        provider.Sources[0].Removed = true;
        bool blocked = false;
        try { provider.DiscoverFrozen(config, scope); } catch (ArchiveReviewException) { blocked = true; }
        Check(blocked, "Identidade ausente entre captura e metadados gerou plano parcial silencioso.");
    }
    private static void OwnTrashMoveIsNotArrival(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"))
        { IncludeTrashInScope = true, ChangeEntryIdOnMove = true };
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root));
        engine.Copy(job); engine.Verify(job); engine.Remove(job, false);
        var rows = engine.Items(job);
        Check(rows.Count == 5 && rows.All(x => x.State == ArchiveItemState.InDeletedItems &&
            x.RemovalEntryId is not null && x.RemovalEntryId != x.SourceEntryId && x.RemovalRecordKey != x.SourceRecordKey),
            "Fixture não exercitou movimentação própria com troca de EntryID e RecordKey.");
        Check(provider.Sources.All(x => !x.Removed && x.Folder == "trash") && provider.DeleteCalls == 0,
            "Modo Itens Excluídos removeu definitivamente mensagens artificiais.");
        var current = engine.CheckNew(job);
        Check(current.Error is null && current.NewItemsCount == 0,
            "Originais movidos pela própria ferramenta à lixeira foram anunciados como novas chegadas.");
        string next = engine.PlanNew(job);
        Check(engine.Items(next).Count == 0, "Plano complementar adotou RemovalEntryIds já catalogados como mensagens novas.");
        engine.Copy(next); engine.Copy(job);
        Check(provider.Copies.Count == 5 && provider.DeleteCalls == 0,
            "Reexecução/plano complementar duplicou originais já arquivados e movidos à lixeira.");
    }
    private static void SameChangeKeySkipsOnlyImportSnapshot(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(provider.SourceMatchesChangeKeyCalls == 5 && provider.SourceSnapshotCalls == 0 && provider.Copies.Count == 5,
            "ChangeKey idêntico não evitou somente o snapshot redundante antes da importação.");
        Check(engine.Items(job).All(x => x.SourceChangeKey == "0" && x.State == ArchiveItemState.Copied), "ChangeKey da exportação não foi preservado no checkpoint.");
        engine.Verify(job);
        Check(provider.SourceSnapshotCalls == 5 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "Otimização por ChangeKey suprimiu a conferência integral da origem na validação.");
    }
    private static void ChangedKeyAndContentBlocksImport(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root));
        engine.Fault = (phase, row) =>
        {
            if (phase != "after-export" || row.SourceEntryId != "source-0") return;
            provider.Sources[0].Version++; provider.Sources[0].Fields["Subject"] = FakeArchiveBackend.Hash("Conteúdo alterado após exportação."); provider.Save();
        };
        engine.Copy(job);
        Check(provider.SourceMatchesChangeKeyCalls == 5 && provider.SourceSnapshotCalls == 1 &&
            engine.Items(job).Single(x => x.SourceEntryId == "source-0").State == ArchiveItemState.ReviewRequired,
            "ChangeKey diferente não exigiu snapshot integral ou não bloqueou conteúdo alterado.");
        Check(provider.Copies.Count == 4 && provider.Copies.All(x => x.SourceId != "source-0") && !provider.Sources[0].Removed,
            "Mensagem alterada após exportação foi importada ou removida.");
    }
    private static void ChangedKeySameContentUsesFullSnapshot(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root));
        engine.Fault = (phase, row) =>
        {
            if (phase == "after-export" && row.SourceEntryId == "source-0") { provider.Sources[0].Version++; provider.Save(); }
        };
        engine.Copy(job);
        Check(provider.SourceMatchesChangeKeyCalls == 5 && provider.SourceSnapshotCalls == 1 && provider.Copies.Count == 5,
            "Mudança apenas de ChangeKey não fez conferência integral ou rejeitou indevidamente campos preservados.");
        engine.Verify(job);
        Check(provider.SourceSnapshotCalls == 6 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "Fallback integral com mesmo conteúdo não manteve a validação integral posterior.");
    }
    private static void MissingChangeKeyUsesFullSnapshot(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog");
        using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json")) { OmitExportChangeKey = true };
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job);
        Check(provider.SourceMatchesChangeKeyCalls == 0 && provider.SourceSnapshotCalls == 5 && provider.Copies.Count == 5 &&
            engine.Items(job).All(x => x.SourceChangeKey is null), "ChangeKey ausente não preservou o caminho de snapshot integral anterior.");
        engine.Verify(job);
        Check(provider.SourceSnapshotCalls == 10 && engine.Items(job).All(x => x.State == ArchiveItemState.Verified),
            "Origem sem ChangeKey perdeu a conferência integral na validação.");
    }
    private static void BlockChangedBackup(string directory)
    {
        Directory.CreateDirectory(directory); string root = Path.Combine(directory, "catalog"); using var provider = new FakeArchiveBackend(Path.Combine(directory, "mail.json"));
        using var engine = new ArchiveEngine(root, provider); string job = engine.Plan(Config(root)); engine.Copy(job); engine.Verify(job);
        var row = engine.Items(job)[0]; File.AppendAllText(row.StageMsgPath!, "corrupt"); engine.Remove(job, true);
        Check(!provider.Sources[0].Removed, "Backup alterado permitiu remoção.");
    }
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    private static void MustFail(Action action) { try { action(); } catch (Exception) { return; } throw new Exception("Operação inválida aceita."); }
}

internal sealed class FakeArchiveBackend : IArchiveBackend, IArchiveDeliveryReportBackend
{
    internal sealed class FakeSource
    {
        public string Id { get; set; } = "";
        public string Folder { get; set; } = "folder";
        public string RecordKey { get; set; } = "";
        public int Version { get; set; }
        public long Size { get; set; } = 800 * 1024;
        public bool Removed { get; set; }
        public DateTime? MessageDateUtc { get; set; }
        public Dictionary<string, string> Fields { get; set; } = new();
    }
    internal sealed record FakeCopy(string JobId, int Volume, string SourceId, string? Token, string EntryId, string FolderId, Dictionary<string, string> Fields, long Size);
    private sealed class State
    {
        public List<FakeSource> Sources { get; set; } = new();
        public List<FakeCopy> Copies { get; set; } = new();
        public Dictionary<string, string> Owners { get; set; } = new();
        public int DeleteCalls { get; set; }
    }
    private readonly string file;
    private State state;
    internal List<FakeSource> Sources => state.Sources;
    internal List<FakeCopy> Copies => state.Copies;
    internal int DeleteCalls => state.DeleteCalls;
    internal bool DropNextToken;
    internal bool ExclusiveDiskRead = true;
    internal bool EmitRelevantMetadata, EmitEmbeddedEvidence;
    internal int FailMoveCount;
    internal bool ThrowOnReopen;
    internal int ReopenCalls;
    internal Action? AfterCapture;
    internal bool IncludeTrashInScope;
    internal bool ChangeEntryIdOnMove;
    internal bool OmitExportChangeKey;
    internal int SourceMatchesChangeKeyCalls;
    internal int SourceSnapshotCalls;
    internal FakeArchiveBackend(string file)
    {
        this.file = file;
        if (File.Exists(file)) state = JsonSerializer.Deserialize<State>(File.ReadAllText(file))!;
        else
        {
            state = new();
            for (int i = 0; i < 5; i++) state.Sources.Add(new() { Id = "source-" + i, RecordKey = "record-" + i,
                Fields = new() { ["MessageClass"] = Hash("IPM.Note"), ["Subject"] = Hash("artificial-" + i),
                    ["ReceivedTime"] = Hash("received"), ["SentOn"] = Hash("sent"), ["AttachmentCount"] = Hash("1"),
                    ["Attachment-1-Hash"] = Hash(Hash("attachment-" + i)),
                    ["NativeBody-MSGBodyStream:__substg1.0_1000001F"] = Hash("length=10;sha256=" + Hash("body-" + i)) } });
            Save();
        }
    }
    internal static string Hash(string text) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    internal void Save()
    {
        string temporary = file + ".tmp"; File.WriteAllText(temporary, JsonSerializer.Serialize(state)); File.Move(temporary, file, true);
    }
    public List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config)
    {
        var folders = Sources.Where(x => !x.Removed && (IncludeTrashInScope || x.Folder != "trash") && (config.IncludeSubfolders || x.Folder == config.SourceFolderId))
            .Select(x => x.Folder).Append(config.SourceFolderId).Distinct(StringComparer.Ordinal).ToList();
        var scope = folders.Select(folder => new FolderScopeSnapshot(folder, ["Pasta artificial"],
            Sources.Where(x => !x.Removed && x.Folder == folder).Select(x => x.Id).ToList())).ToList();
        var hook = AfterCapture; AfterCapture = null; hook?.Invoke(); return scope;
    }
    public List<SourceDescriptor> DiscoverFrozen(ArchiveConfig config, IReadOnlyList<FolderScopeSnapshot> scope)
    {
        var rows = new List<SourceDescriptor>();
        foreach (var folder in scope)
            foreach (string id in folder.EntryIds)
            {
                var source = Sources.SingleOrDefault(x => x.Id == id && !x.Removed && x.Folder == folder.FolderId)
                    ?? throw new ArchiveReviewException("Mensagem congelada não está mais na pasta original.");
                if (!ArchiveDateFilter.Includes(source.MessageDateUtc, config.CutoffDateInclusive, config.CutoffTimeZoneId)) continue;
                rows.Add(new("source", folder.FolderId, source.Id, folder.RelativePath, source.Size, true, true, source.RecordKey,
                    MessageDateUtc: config.CutoffDateInclusive is null ? null : source.MessageDateUtc,
                    DateBasis: config.CutoffDateInclusive is null ? null : "received", MessageClass: source.Fields.GetValueOrDefault("MessageClass")==Hash(DeliveryReportEvidence.MessageClass)?DeliveryReportEvidence.MessageClass:null));
            }
        return rows;
    }
    public SourceDescriptor? InspectDeliveryReport(ArchiveItem item)
    {
        var source=Sources.SingleOrDefault(x=>x.Id==item.SourceEntryId && !x.Removed && x.Folder==item.SourceFolderId);
        if(source is not null)
            foreach(string messageClass in new[]{"IPM.Schedule.Meeting.Request","IPM.Schedule.Meeting.Resp.Pos","IPM.Schedule.Meeting.Canceled"})
                if(source.Fields.GetValueOrDefault("MessageClass")==Hash(messageClass))return new("source",source.Folder,source.Id,item.RelativePath,source.Size,false,false,source.RecordKey,MessageClass:messageClass);
        return source is not null && source.Fields.GetValueOrDefault("MessageClass")==Hash(DeliveryReportEvidence.MessageClass)
            ? new("source",source.Folder,source.Id,item.RelativePath,source.Size,true,true,source.RecordKey,source.MessageDateUtc,"received",DeliveryReportEvidence.MessageClass):null;
    }
    public List<SourceDescriptor> Discover(ArchiveConfig config) => DiscoverFrozen(config, CaptureScope(config));
    public ArchiveSnapshot Export(ArchiveItem item, string evidenceDirectory) => Capture(Source(item.SourceEntryId).Fields, evidenceDirectory,
        OmitExportChangeKey ? null : Source(item.SourceEntryId).Version.ToString());
    public bool SourceMatchesChangeKey(ArchiveItem item, string expected)
    {
        SourceMatchesChangeKeyCalls++;
        var source = Source(item.SourceEntryId);
        if (source.Folder != item.SourceFolderId || item.SourceRecordKey is not null && source.RecordKey != item.SourceRecordKey)
            throw new ArchiveReviewException("Origem fora da pasta ou com identidade alterada.");
        return source.Version.ToString() == expected;
    }
    public ArchiveSnapshot SourceSnapshot(ArchiveItem item, SourceLocation? location, string evidenceDirectory)
    {
        SourceSnapshotCalls++;
        var source = Source(location?.EntryId ?? item.SourceEntryId);
        if (source.Folder != (location?.FolderId ?? item.SourceFolderId)) throw new ArchiveReviewException("Origem fora da pasta.");
        return Capture(source.Fields, evidenceDirectory, source.Version.ToString());
    }
    public ArchiveSnapshot BackupSnapshot(ArchiveItem item, string evidenceDirectory)
    {
        var fields = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(item.StageMsgPath!))!;
        Directory.CreateDirectory(evidenceDirectory); ArchiveFiles.WriteNew(Path.Combine(evidenceDirectory, "field-hashes.json"), fields);
        return WithMetadata(new(fields, item.StageMsgPath!, ArchiveFiles.Sha(item.StageMsgPath!)));
    }
    private ArchiveSnapshot Capture(Dictionary<string, string> fields, string directory, string? changeKey = null)
    {
        Directory.CreateDirectory(directory); string path = Path.Combine(directory, "message.msg"); ArchiveFiles.WriteNew(path, fields);
        ArchiveFiles.WriteNew(Path.Combine(directory, "field-hashes.json"), fields);
        return WithMetadata(new(new(fields), path, ArchiveFiles.Sha(path), changeKey));
    }
    private ArchiveSnapshot WithMetadata(ArchiveSnapshot snapshot)
    {
        if (!EmitRelevantMetadata) return snapshot;
        var profile = new Dictionary<string,string> { ["ProfileVersion"] = "relevant-message-v1", ["BodyMode"] = "html-exact-v1", ["Body-HTML"] = "same-body" };
        foreach (string key in new[] { "SenderName", "SenderEmailType", "SenderEmailAddress", "SentOnBehalfOfName", "SentUtc", "ReceivedUtc", "Recipients" }) profile["Meta-" + key] = "same";
        var embedded = EmbeddedMessageEvidence.Indices(snapshot.Fields);
        if (EmitEmbeddedEvidence && embedded.Count > 0)
        {
            profile["Embedded-Version"] = "embedded-msg-v1"; profile["Embedded-Indices"] = string.Join(",", embedded);
            foreach (int index in embedded) profile["Embedded-Item-" + index] = Hash("same nested message " + index);
        }
        return snapshot with { RelevantFields = profile };
    }
    public void CreateVolume(ArchiveVolume volume)
    {
        if (volume.IsExisting)
        {
            ArchiveFiles.ValidateExistingPst(volume.PstPath);
            state.Owners[OwnerKey(volume)] = volume.JobId + ":" + volume.Number;
            volume.StoreId = "target-" + volume.JobId + "-" + volume.Number; Save(); return;
        }
        if (File.Exists(volume.PstPath)) { AttachVolume(volume); return; }
        using (var f = new FileStream(volume.PstPath, FileMode.CreateNew)) f.SetLength(ArchiveVolumePolicy.DefaultEmptyVolumeBaseBytes);
        state.Owners[OwnerKey(volume)] = volume.JobId + ":" + volume.Number; volume.StoreId = "target-" + volume.JobId + "-" + volume.Number; Save();
    }
    public void AttachVolume(ArchiveVolume volume)
    {
        if (!File.Exists(volume.PstPath) || state.Owners.GetValueOrDefault(OwnerKey(volume)) != volume.JobId + ":" + volume.Number) throw new ArchiveReviewException("PST estranho.");
        volume.StoreId = "target-" + volume.JobId + "-" + volume.Number;
    }
    private static string OwnerKey(ArchiveVolume volume) => volume.IsExisting ? volume.PstPath + "|" + volume.JobId : volume.PstPath;
    public long MeasureVolume(ArchiveVolume volume) => new FileInfo(volume.PstPath).Length;
    public TokenLookup FindToken(ArchiveVolume volume, ArchiveItem item)
    {
        AttachVolume(volume); var copies = Copies.Where(x => x.JobId == volume.JobId && x.Volume == volume.Number).ToList();
        return new(copies.Where(x => x.Token == item.DestinationToken).Select(x => new DestinationIdentity(volume.StoreId!, x.FolderId, x.EntryId)).ToList(), copies.Count(x => x.Token is null));
    }
    public DestinationIdentity Import(ArchiveVolume volume, ArchiveItem item)
    {
        var source = Sources.Single(x => x.Id == item.SourceEntryId);
        var stagedFields = JsonSerializer.Deserialize<Dictionary<string,string>>(File.ReadAllText(item.StageMsgPath!))!;
        var copy = new FakeCopy(volume.JobId, volume.Number, source.Id, DropNextToken ? null : item.DestinationToken,
            Guid.NewGuid().ToString("N"), "target-folder", new(stagedFields), source.Size); DropNextToken = false;
        Copies.Add(copy); using (var f = new FileStream(volume.PstPath, FileMode.Open)) f.SetLength(f.Length + copy.Size); Save();
        return new(volume.StoreId!, copy.FolderId, copy.EntryId);
    }
    public ArchiveSnapshot TargetSnapshot(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity, string evidenceDirectory, bool requireToken = true)
    {
        var copy = Copies.Single(x => x.EntryId == identity.EntryId && (!requireToken || x.Token == item.DestinationToken) && x.JobId == volume.JobId);
        return Capture(copy.Fields, evidenceDirectory);
    }
    public List<DestinationIdentity> UnmarkedCandidates(ArchiveVolume volume, ArchiveItem item) => Copies.Where(x => x.JobId == volume.JobId && x.Volume == volume.Number && x.Token is null)
        .Select(x => new DestinationIdentity(volume.StoreId!, x.FolderId, x.EntryId)).ToList();
    public void MarkCopy(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity)
    {
        int index = Copies.FindIndex(x => x.EntryId == identity.EntryId && x.JobId == volume.JobId);
        if (index < 0 || Copies[index].Token is not null && Copies[index].Token != item.DestinationToken) throw new ArchiveReviewException("Candidato já identificado.");
        Copies[index] = Copies[index] with { Token = item.DestinationToken }; Save();
    }
    public DiskRoundtrip ReopenVolume(ArchiveVolume volume)
    {
        ReopenCalls++;
        if (ThrowOnReopen) throw new IOException("Falha transitória artificial de reabertura.");
        AttachVolume(volume); return new(true, ExclusiveDiskRead);
    }
    public SourceLocation? LocateSource(ArchiveItem item)
    {
        var source = Sources.SingleOrDefault(x => x.Id == (item.RemovalEntryId ?? item.SourceEntryId) && !x.Removed);
        return source is null ? null : new(source.Id, source.Folder, source.Folder == "trash", source.RecordKey);
    }
    public SourceLocation MoveToDeletedItems(ArchiveItem item, string? expectedChangeKey)
    {
        if (FailMoveCount > 0) { FailMoveCount--; throw new System.Runtime.InteropServices.COMException("Falha artificial do Outlook.", unchecked((int)0x80004005)); }
        var source = Source(item.SourceEntryId); CheckVersion(source, expectedChangeKey); source.Folder = "trash"; source.Version++;
        if (ChangeEntryIdOnMove)
        {
            source.Id = "trash-" + Guid.NewGuid().ToString("N"); source.RecordKey = "record-" + Guid.NewGuid().ToString("N");
        }
        Save();
        return new(source.Id, source.Folder, true, source.RecordKey);
    }
    public void DeleteFromDeletedItems(ArchiveItem item, SourceLocation location, string? expectedChangeKey)
    {
        var source = Source(location.EntryId); CheckVersion(source, expectedChangeKey);
        if (source.Folder != "trash") throw new ArchiveReviewException("Não está na lixeira."); source.Removed = true; state.DeleteCalls++; Save();
    }
    private FakeSource Source(string id) => Sources.SingleOrDefault(x => x.Id == id && !x.Removed) ?? throw new ArchiveReviewException("Origem ausente.");
    private static void CheckVersion(FakeSource source, string? expected) { if (source.Version.ToString() != expected) throw new ArchiveReviewException("Origem mudou."); }
    public void Dispose() { }
}
