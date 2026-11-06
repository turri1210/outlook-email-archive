internal static class ArchiveImmediateTests
{
    internal static int Run()
    {
        foreach(string mode in new[]{"fresh","missing-key","resume","different-target","changed-source","tampered-backup","without-consent"})
        {
            string path=Path.Combine(Path.GetTempPath(),"arc-immediate-"+Guid.NewGuid().ToString("N")),root=Path.Combine(path,"catalog");Directory.CreateDirectory(path);
            using var backend=new FakeArchiveBackend(Path.Combine(path,"provider.json")){ExclusiveDiskRead=false,OmitExportChangeKey=mode=="missing-key"};
            using var engine=new ArchiveEngine(root,backend);string job=engine.Plan(new ArchiveConfig{RootDirectory=root,SourceStoreId="source",SourceFolderId="folder"});
            if(mode is "different-target" or "changed-source" or "tampered-backup")
            {
                engine.Fault=(phase,row)=>{
                    if(phase!="after-import-before-checkpoint")return;
                    if(mode=="different-target")backend.Copies.Single(x=>x.SourceId==row.SourceEntryId).Fields["Subject"]=FakeArchiveBackend.Hash("changed target");
                    if(mode=="changed-source"){var original=backend.Sources.Single(x=>x.Id==row.SourceEntryId);original.Fields["Subject"]=FakeArchiveBackend.Hash("changed source");original.Version++;}
                    if(mode=="tampered-backup")File.AppendAllText(row.StageMsgPath!,"corruption");
                    backend.Save();
                };
            }
            if(mode=="resume")
            {
                bool injected=false;engine.Fault=(phase,_)=>{if(phase=="after-trash-move-before-checkpoint"&&!injected){injected=true;throw new ArchiveInjectedCrashException("immediate");}};
                try{engine.Copy(job,true,true);throw new Exception("Interrupção não ocorreu.");}catch(ArchiveInjectedCrashException){}
                engine.Fault=null;
                Check(engine.ContinueWorkflow(job,true,allowRetainedBackup:true)==0,"Retomada não concluiu a movimentação interrompida.");
            }
            else engine.Copy(job,true,mode!="without-consent");
            var rows=engine.Items(job);
            if(mode is "fresh" or "missing-key" or "resume")
            {
                Check(rows.All(x=>x.State==ArchiveItemState.InDeletedItems) && backend.Copies.Count==5,"Fluxo imediato não concluiu ou duplicou cópias.");
                if(mode=="fresh")Check(backend.SourceSnapshotCalls==0 && backend.ReopenCalls==0,"Fluxo imediato releu o original ou reabriu o PST desnecessariamente.");
                if(mode=="missing-key")Check(backend.SourceSnapshotCalls>=5,"ChangeKey ausente pulou a leitura integral.");
            }
            else Check(backend.Sources.All(x=>x.Folder=="folder"),"Originais foram movidos sem consentimento ou com divergências.");
            Check(backend.DeleteCalls==0,"Fluxo imediato fez exclusão definitiva.");
            Console.WriteLine("Fluxo imediato: "+mode+" aprovado.");
        }
        ArchiveValidation.SelfTest();return 0;
    }
    private static void Check(bool value,string error){if(!value)throw new Exception(error);}
}
