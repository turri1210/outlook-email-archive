using System.Text.Json;

internal sealed partial class ArchiveEngine
{
    internal void ProfileCopySample(string jobId,int count)
    {
        if(count is <1 or >5)throw new ArgumentException("Use de uma a cinco mensagens para a amostra.");
        var job=ledger.GetJob(jobId);WorkflowPlan.Ensure(ledger,job,AutomaticMovementEnabled(jobId));
        var selection=ledger.GetItems(jobId).Where(x=>x.State==ArchiveItemState.Planned && x.NativeBatchId is null && (x.SourceMessageClass is null or "IPM.Note") && x.EstimatedBytes<2*1024*1024).Take(count).Select(x=>x.Key).ToHashSet(StringComparer.Ordinal);
        if(selection.Count==0)throw new ArchiveReviewException("Nenhuma mensagem elegível para medição.");
        if(workflowSelection is not null)throw new InvalidOperationException("Outro escopo já está ativo.");
        workflowSelection=selection;
        try{using var timing=ArchiveOperationTiming.Measure("copy.sample");Copy(jobId);}
        finally{workflowSelection=null;}
    }
}

internal static class ArchivePerformanceAudit
{
    internal static int SelfTest()
    {
        string path=Path.Combine(Path.GetTempPath(),"archive-timing-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(path);
        string log=Path.Combine(path,"timings.jsonl");
        using(ArchiveOperationTiming.Measure("disabled")){}
        using(ArchiveOperationTiming.Begin(log))
        using(var backend=new FakeArchiveBackend(Path.Combine(path,"provider.json")))
        using(var engine=new ArchiveEngine(Path.Combine(path,"catalog"),backend))
        {
            string job=engine.Plan(new ArchiveConfig{RootDirectory=Path.Combine(path,"catalog"),SourceStoreId="source",SourceFolderId="folder",MoveValidatedToDeletedItems=true});
            engine.ProfileCopySample(job,3);
            if(backend.Copies.Count!=3 || backend.DeleteCalls!=0 || backend.Sources.Any(x=>x.Folder!="folder"))throw new Exception("Amostra de desempenho alterou originais ou saiu do limite.");
            if(engine.Items(job).Count(x=>x.State==ArchiveItemState.Copied)!=3)throw new Exception("Amostra não foi registrada para retomada.");
        }
        var rows=File.ReadLines(log).Select(line=>JsonSerializer.Deserialize<JsonElement>(line)).ToList();
        var ids=rows.Select(x=>x.GetProperty("Id").GetInt64()).ToHashSet();
        if(ids.Count!=rows.Count || rows.Any(x=>x.GetProperty("Milliseconds").GetDouble()<0 || x.GetProperty("Parent").ValueKind!=JsonValueKind.Null&&!ids.Contains(x.GetProperty("Parent").GetInt64())))throw new Exception("Hierarquia das medições inválida.");
        Console.WriteLine("Desempenho: hierarquia das medições e três cópias sintéticas sem movimentar originais aprovadas.");return 0;
    }
    internal static int Run(string root,string jobId,string output)
    {
        ArchiveFiles.ValidateLocalDirectory(output);if(Directory.Exists(output))throw new IOException("Use uma pasta nova para o diagnóstico.");Directory.CreateDirectory(output);
        using var timing=ArchiveOperationTiming.Begin(Path.Combine(output,"timings.jsonl"));
        using(var backend=new OutlookArchiveBackend())
        using(var engine=new ArchiveEngine(root,backend))engine.ProfileCopySample(jobId,3);
        using var operationLock=new FileStream(Path.Combine(root,"archive.lock"),FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None);
        using var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);
        using var reader=new OutlookArchiveBackend{RetainComparisonSnapshots=true};
        var candidates=ledger.GetItems(jobId).Where(x=>x.State==ArchiveItemState.ReviewRequired && x.StageMsgPath is not null && x.TargetEntryId is not null).ToList();
        var sample=candidates.Where((_,i)=>i%Math.Max(1,candidates.Count/12)==0).Take(12).ToList();
        var rows=new List<object>();
        foreach(var item in sample)
        {
            using var itemTiming=ArchiveOperationTiming.Measure("validation.sample",item.Key[..12]);
            string folder=Path.Combine(output,item.Key[..12]);Directory.CreateDirectory(folder);
            try
            {
                using var lease=new ArchiveEvidenceLease(item);
                var volume=ledger.GetVolumes(jobId).Single(v=>v.Number==item.VolumeNumber);
                using(ArchiveOperationTiming.Measure("validation.locate-source")){if(reader.LocateSource(item) is null)throw new ArchiveReviewException("Origem não encontrada.");}
                ArchiveSnapshot source,backup,target;
                using(ArchiveOperationTiming.Measure("validation.source"))source=reader.Export(item,Path.Combine(folder,"source"));
                using(ArchiveOperationTiming.Measure("validation.backup"))backup=reader.BackupSnapshot(item,Path.Combine(folder,"backup"));
                using(ArchiveOperationTiming.Measure("validation.target"))target=reader.TargetSnapshot(volume,item,new(item.TargetStoreId!,item.TargetFolderId!,item.TargetEntryId!),Path.Combine(folder,"target"));
                bool equivalent;using(ArchiveOperationTiming.Measure("validation.compare"))equivalent=RelevantMessageValidation.Equivalent(backup,target);
                rows.Add(new{Key=item.Key[..12],Equivalent=equivalent,Source=source.MessagePath,Backup=backup.MessagePath,Target=target.MessagePath});
            }
            catch(Exception ex){rows.Add(new{Key=item.Key[..12],Error=ex.GetType().Name+": "+ex.Message});}
        }
        ArchiveFiles.WriteNew(Path.Combine(output,"samples.json"),rows);
        Console.WriteLine("Diagnóstico concluído: três cópias da execução medidas; originais não movidos pela medição. Amostras de validação somente lidas.");return 0;
    }
}
