using System.Text.Json;

internal static class ArchiveWorkflowTests
{
    internal static int Run()
    {
        string basePath=Path.Combine(Path.GetTempPath(),"arc-flow-"+Guid.NewGuid().ToString("N")[..12]);Directory.CreateDirectory(basePath);
        foreach(string scenario in new[]{"pause","crash","review","migration","mode","tamper","checkpoint","blocked"})
        {
            string path=Path.Combine(basePath,scenario),root=Path.Combine(path,"catalog");Directory.CreateDirectory(path);
            using var backend=new FakeArchiveBackend(Path.Combine(path,"mail.json"));
            var original=backend.Sources[0];backend.Sources.Clear();
            for(int i=0;i<205;i++)backend.Sources.Add(new(){Id="source-"+i,RecordKey="record-"+i,Size=16*1024,Fields=new(original.Fields)});
            backend.Save();using var engine=new ArchiveEngine(root,backend);
            var config=new ArchiveConfig{RootDirectory=root,SourceStoreId="source",SourceFolderId="folder",TargetVolumeBytes=128*ArchiveVolumePolicy.MiB};
            string job=engine.Plan(config);
            if(scenario=="blocked")
            {
                backend.ExclusiveDiskRead=false;
                try{engine.ContinueWorkflow(job,true);throw new Exception("PST ocupado foi tratado como conclusão ou pausa.");}
                catch(ArchiveReviewException ex){Check(ex.Message.Contains("PST ocupado"),"Erro de bloqueio perdeu a causa.");}
                using(var blocked=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true))Check(blocked.WorkflowStatus(job)!.State=="Blocked","Bloqueio foi registrado como pausa voluntária.");
                Check(backend.Copies.Count==100 && backend.Sources.All(x=>x.Folder=="folder"),"Falha de validação abriu outro lote ou moveu originais.");
                backend.ExclusiveDiskRead=true;
                Check(engine.ContinueWorkflow(job,true)==0 && backend.Copies.Count==205,"Retomada após bloqueio duplicou cópias ou não concluiu.");
            }
            else if(scenario=="migration")
            {
                engine.Copy(job);engine.Verify(job,moveValidatedToDeletedItems:true);
                int reads=backend.SourceSnapshotCalls;Check(engine.ContinueWorkflow(job,true)==0,"Migração de execução concluída falhou.");
                Check(backend.SourceSnapshotCalls==reads,"Migração reabriu mensagens já concluídas.");
            }
            else if(scenario=="mode")
            {
                Check(engine.ContinueWorkflow(job,false)==0 && backend.Sources.All(x=>x.Folder=="folder"),"Modo de manter originais foi desrespeitado.");
                Check(engine.ContinueWorkflow(job,true)==0 && backend.Sources.All(x=>x.Folder=="trash"),"Mudança de objetivo não reabriu os lotes necessários.");
            }
            else if(scenario=="review")
            {
                engine.Copy(job);backend.Sources[0].Fields["Subject"]=FakeArchiveBackend.Hash("changed");backend.Save();
                Check(engine.ContinueWorkflow(job,true)==3,"Divergência não foi separada.");
                Check(backend.Sources.Count(x=>x.Folder=="trash")==204 && backend.Sources[0].Folder=="folder","Revisão bloqueou lotes posteriores ou removeu divergência.");
                int reads=backend.SourceSnapshotCalls;engine.ContinueWorkflow(job,true);Check(reads==backend.SourceSnapshotCalls,"Revisão foi repetida sem solicitação.");
                Check(engine.ContinueWorkflow(job,true,reconcilePending:true)==3,"Reconciliação aceitou divergência ainda presente.");
                Check(backend.Sources[0].Folder=="folder" && backend.Copies.Count==205,"Reconciliação alterou origem divergente ou duplicou cópias.");
                // Advanced actions or interrupted status writes can leave an
                // outdated checkpoint. Actual item states must drive Continue.
                using(var repairLedger=new ArchiveLedger(Path.Combine(root,"archive.sqlite")))repairLedger.SetWorkflowState(job,1,"Complete");
                backend.Sources[0].Fields["Subject"]=original.Fields["Subject"];backend.Save();
                Check(engine.ContinueWorkflow(job,true,reconcilePending:true)==0,"Continuar não reconciliou e concluiu a pendência recuperada.");
                Check(backend.Sources.All(x=>x.Folder=="trash") && backend.Copies.Count==205,"Recuperação não moveu somente os originais validados.");
            }
            else
            {
                int checks=0;
                if(scenario is "crash" or "checkpoint")
                {
                    string phaseName=scenario=="crash"?"after-trash-move-before-checkpoint":"after-workflow-batch-checkpoint";
                    bool interrupted=false;engine.Fault=(phase,_)=>{if(phase==phaseName&&!interrupted){interrupted=true;throw new ArchiveInjectedCrashException(phase);}};
                    try{engine.ContinueWorkflow(job,true);throw new Exception("Interrupção não ocorreu.");}catch(ArchiveInjectedCrashException){}
                    engine.Fault=null;
                }
                else
                {
                    if(scenario=="pause")engine.Fault=(phase,_)=>{if(phase=="after-workflow-batch-checkpoint")File.WriteAllText(Path.Combine(root,"jobs",job,"workflow","pause-after-batch.request"),"pause");};
                    Check(engine.ContinueWorkflow(job,true,pauseRequested:scenario=="pause"?null:()=>++checks>=2)==4,"Pausa ao concluir lote não ocorreu.");
                    engine.Fault=null;
                    Check(backend.Copies.Count==100 && backend.Sources.Count(x=>x.Folder=="trash")==100,"O fluxo processou itens de outro lote antes da pausa.");
                }
                if(scenario=="tamper")
                {
                    string file=Path.Combine(root,"jobs",job,"workflow","manifest.json");File.AppendAllText(file," ");
                    int count=backend.SourceSnapshotCalls;try{engine.ContinueWorkflow(job,true);throw new Exception("Manifesto adulterado aceito.");}catch(ArchiveReviewException){}
                    Check(count==backend.SourceSnapshotCalls,"Adulteração abriu mensagens antes de bloquear.");continue;
                }
                int before=backend.SourceSnapshotCalls;Check(engine.ContinueWorkflow(job,true)==0,"Retomada não concluiu.");
                if(scenario is "pause" or "checkpoint")Check(backend.SourceSnapshotCalls-before==105,"Lote concluído foi verificado novamente na retomada.");
                Check(backend.Copies.Count==205,"Retomada duplicou cópias.");
            }
            Check(backend.DeleteCalls==0,"O fluxo fez exclusão definitiva.");
            int completedReads=backend.SourceSnapshotCalls;engine.ContinueWorkflow(job,true);Check(completedReads==backend.SourceSnapshotCalls,"Fluxo finalizado releu mensagens.");
            using var read=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);var progress=read.WorkflowStatus(job)!;
            Check(progress.TotalBatches==3 && progress.TotalItems==205,"Escopo de lotes incorreto.");
            Console.WriteLine("Fluxo por lotes: "+scenario+" aprovado.");
        }
        Console.WriteLine("Fluxo por lotes: oito cenários aprovados, sem Outlook ou mensagens reais.");return 0;
    }
    private static void Check(bool value,string error){if(!value)throw new Exception(error);}
}
