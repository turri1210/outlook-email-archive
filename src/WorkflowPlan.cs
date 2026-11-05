using System.Text.Json;

internal static class WorkflowPlan
{
    internal sealed record Manifest(string Version,string Job,string PlanSha,List<List<string>> Batches);
    internal static bool Done(ArchiveItemState state,bool move)=>state is ArchiveItemState.InDeletedItems or ArchiveItemState.Removed or ArchiveItemState.BackupVerifiedSourceUnmanaged or ArchiveItemState.SkippedByPolicy || !move&&state==ArchiveItemState.Verified;
    internal static bool Actionable(ArchiveItemState state,bool move)=>state is ArchiveItemState.Planned or ArchiveItemState.Exported or ArchiveItemState.ImportStarted or ArchiveItemState.Copied or ArchiveItemState.DeleteStarted || move&&state==ArchiveItemState.Verified;
    internal static string Classify(IEnumerable<ArchiveItem> items,bool move)
    {var states=items.Select(x=>x.State).ToList();return states.Any(x=>Actionable(x,move))?"Pending":states.All(x=>Done(x,move))?"Complete":"Review";}
    internal static void Ensure(ArchiveLedger ledger,ArchiveJob job,bool move,bool recoverTransient=false)
    {
        var config=JsonSerializer.Deserialize<ArchivePlanEnvelope>(job.ConfigJson)!;
        string plan=ArchiveFiles.Child(job.RootDirectory,Path.Combine(job.RootDirectory,"source-plan.json"));
        if(ArchiveFiles.Sha(plan)!=config.PlanSha256)throw new ArchiveReviewException("Plano original alterado; lotes não serão preparados.");
        void RequeueTransient()
        {
            if(!recoverTransient)return;
            int queued=ledger.RequeueUnavailableNativeIdentity(job.Id);
            if(queued>0)Console.WriteLine($"Conversão: {queued:N0} cópias reenfileiradas para validação completa; nenhuma aprovação ou remoção concedida.");
        }
        if(ledger.WorkflowManifest(job.Id) is { } known)
        {
            string file=ArchiveFiles.Child(job.RootDirectory,known.Path);
            if(ArchiveFiles.Sha(file)!=known.Sha)throw new ArchiveReviewException("Manifesto dos lotes alterado.");
            var retained=JsonSerializer.Deserialize<Manifest>(File.ReadAllText(file))!;
            var catalog=ledger.WorkflowBatches(job.Id);
            if(retained.Version!="archive-workflow-v1" || retained.Job!=job.Id || retained.PlanSha!=config.PlanSha256 || JsonSerializer.Serialize(retained.Batches)!=JsonSerializer.Serialize(catalog.Select(x=>x.Items).ToList()))
                throw new ArchiveReviewException("Lotes do catálogo não correspondem ao manifesto preservado.");
            RequeueTransient();
            ledger.SetWorkflowMovement(job.Id,move);return;
        }
        RequeueTransient();
        var rows=ledger.GetItems(job.Id);
        var frozen=JsonSerializer.Deserialize<List<SourceDescriptor>>(File.ReadAllText(plan))!.Select(x=>ArchiveLedger.ComputeItemKey(x.StoreId,x.EntryId)).ToHashSet(StringComparer.Ordinal);
        if(rows.Any(x=>!frozen.Contains(x.Key)))throw new ArchiveReviewException("Item fora do plano inicial; conversão bloqueada.");
        var batches=new List<List<string>>();var current=new List<string>();
        // Keep every existing native-copy group together. A legacy partial copy
        // must never be split across two workflow checkpoints.
        foreach(var unit in rows.GroupBy(x=>x.NativeBatchId??"item:"+x.Key))
        {
            var keys=unit.Select(x=>x.Key).ToList();if(keys.Count>100)throw new ArchiveReviewException("Lote nativo excede o limite esperado.");
            if(current.Count>0&&current.Count+keys.Count>100){batches.Add(current);current=[];}
            current.AddRange(keys);
        }
        if(current.Count>0)batches.Add(current);
        string directory=Path.Combine(job.RootDirectory,"workflow");Directory.CreateDirectory(directory);string manifest=Path.Combine(directory,"manifest.json");
        var data=new Manifest("archive-workflow-v1",job.Id,config.PlanSha256,batches);
        if(File.Exists(manifest))
        {
            var retained=JsonSerializer.Deserialize<Manifest>(File.ReadAllText(manifest))!;
            if(JsonSerializer.Serialize(retained)!=JsonSerializer.Serialize(data))throw new ArchiveReviewException("Manifesto interrompido não corresponde ao catálogo atual.");
        }
        else ArchiveFiles.WriteNew(manifest,data);
        var byKey=rows.ToDictionary(x=>x.Key);
        ledger.AddWorkflow(job.Id,manifest,ArchiveFiles.Sha(manifest),move,batches.Select((keys,index)=>new WorkflowBatch(index+1,keys,Classify(keys.Select(k=>byKey[k]),move))).ToList());
    }
}
