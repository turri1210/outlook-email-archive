internal static class ArchivePlanningTests
{
    internal static int Run()
    {
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
