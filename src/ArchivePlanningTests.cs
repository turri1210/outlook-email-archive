internal static class ArchivePlanningTests
{
    internal static int Run()
    {
        int attempts=0,waits=0;
        int result=ArchivePlanningConnection.Read(()=>{attempts++;if(attempts<3)throw new System.Runtime.InteropServices.COMException("Network problems",unchecked((int)0x80040115));return 7;},_=>waits++);
        Check(result==7 && attempts==3 && waits==2,"Falha transitória não repetiu somente a leitura pendente.");
        attempts=0;
        try{ArchivePlanningConnection.Read<int>(()=>{attempts++;throw new System.Runtime.InteropServices.COMException("Network problems",unchecked((int)0x80040115));},_=>{});throw new Exception("Falha persistente não foi interrompida.");}
        catch(ArchiveConnectionException){Check(attempts==3,"Número de tentativas sem limite.");}
        attempts=0;
        try{ArchivePlanningConnection.Read<int>(()=>{attempts++;throw new System.Runtime.InteropServices.COMException("Item missing",unchecked((int)0x8004010F));},_=>{});throw new Exception("Item ausente foi ignorado.");}
        catch(System.Runtime.InteropServices.COMException){Check(attempts==1,"Item ausente foi tratado como falha de rede.");}
        foreach(string phase in new[]{"scope-checkpoint","item-checkpoint","before-catalog"})
        {
            string path=Path.Combine(Path.GetTempPath(),"archive-planning-"+Guid.NewGuid().ToString("N"));
            string root=Path.Combine(path,"catalog"),provider=Path.Combine(path,"provider.json");Directory.CreateDirectory(path);
            string job;int checkedCount;
            using(var backend=new FakeArchiveBackend(provider))
            using(var engine=new ArchiveEngine(root,backend))
            {
                engine.PlanningFault=(point,count)=>{if(point==phase && (phase!="item-checkpoint" || count==3))throw new ArchiveInjectedCrashException("planning");};
                try{engine.Plan(new ArchiveConfig{RootDirectory=root,SourceStoreId="source",SourceFolderId="folder"});throw new Exception("Interrupção não ocorreu.");}
                catch(ArchiveInjectedCrashException){}
                using var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);var pending=ledger.GetJobs().Single();job=pending.Id;
                Check(pending.PlanningPending,"Planejamento interrompido não consta no histórico.");checkedCount=ledger.PlanningCount(job);
                Check(checkedCount==(phase=="scope-checkpoint"?0:phase=="item-checkpoint"?3:5),"Checkpoint salvo na posição errada.");
                try{engine.Copy(job);throw new Exception("Cópia de plano incompleto autorizada.");}catch(ArchiveReviewException){}
                // A completed checkpoint must not be read from Outlook again.
                foreach(var source in backend.Sources.Take(checkedCount))source.Removed=true;
                var template=backend.Sources[0];backend.Sources.Add(new(){Id="arrival-after-interruption",RecordKey="new-record",Fields=new(template.Fields),Size=template.Size});backend.Save();
            }
            using(var backend=new FakeArchiveBackend(provider))
            using(var engine=new ArchiveEngine(root,backend))
            {
                Check(engine.ResumePlanning(job)==job,"Retomada criou outra execução.");
                var items=engine.Items(job);Check(items.Count==5 && items.All(x=>x.SourceEntryId!="arrival-after-interruption"),"Retomada alterou o escopo inicial ou duplicou itens.");
                using var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);Check(!ledger.GetJob(job).PlanningPending && ledger.PlanningCount(job)==5,"Planejamento não foi concluído.");
                Check(backend.Copies.Count==0 && backend.DeleteCalls==0,"Planejamento copiou ou excluiu mensagens.");
            }
            Console.WriteLine("Retomada do planejamento: "+phase+" aprovada.");
        }
        return 0;
    }
    private static void Check(bool value,string message){if(!value)throw new Exception(message);}
}
