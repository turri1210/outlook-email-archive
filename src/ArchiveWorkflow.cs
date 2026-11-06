using System.Diagnostics;
using System.Text.Json;

internal sealed partial class ArchiveEngine
{
    private HashSet<string>? workflowSelection;
    private Action? workflowProgress;
    private List<ArchiveItem> WorkItems(string job)=>workflowSelection is null?ledger.GetItems(job):ledger.WorkflowItems(job,workflowSelection);

    internal int ContinueWorkflow(string jobId,bool move,bool allowRetainedBackup=false,Func<bool>? pauseRequested=null,bool reconcilePending=false)
    {
        var job=ledger.GetJob(jobId);var config=Envelope(job).Config;
        WorkflowPlan.Ensure(ledger,job,move);
        if(!NativeWorkersIdle(jobId))throw new ArchiveReviewException("Uma cópia nativa anterior ainda está em andamento. Aguarde sua conclusão antes de continuar; o histórico foi preservado.");
        var batches=ledger.WorkflowBatches(jobId);
        // Validate catalog-only ownership once. No concluded message is opened.
        var assignment=batches.SelectMany(b=>b.Items.Select(k=>(Key:k,Number:b.Number))).ToDictionary(x=>x.Key,x=>x.Number);
        var owned=ledger.GetItems(jobId);
        if(owned.Count!=assignment.Count || owned.Any(x=>!assignment.ContainsKey(x.Key)))throw new ArchiveReviewException("Escopo dos lotes não corresponde aos itens da execução.");
        foreach(var group in owned.Where(x=>x.NativeBatchId is not null).GroupBy(x=>x.NativeBatchId))
            if(group.Select(x=>assignment[x.Key]).Distinct().Count()!=1)throw new ArchiveReviewException("Lote nativo dividido entre checkpoints; operação bloqueada.");
        if(backend is OutlookArchiveBackend outlook)outlook.CheckWorkflowReady(config,ledger.GetVolumes(jobId));
        string request=Path.Combine(job.RootDirectory,"workflow","pause-after-batch.request");
        if(File.Exists(request))File.Delete(request); // A deliberate Continue acknowledges an earlier pause.
        var watch=Stopwatch.StartNew();var initial=ledger.WorkflowStatus(jobId)!;long lastReport=-1000;
        int current=0;
        void Report(string activity,bool force=false)
        {
            if(!force && watch.ElapsedMilliseconds-lastReport<1000)return;lastReport=watch.ElapsedMilliseconds;
            var progress=ledger.WorkflowStatus(jobId)!;progress.CurrentBatch=current;progress.Activity=activity;
            int processed=progress.CompleteItems+progress.ReviewItems-initial.CompleteItems-initial.ReviewItems;
            progress.ItemsPerSecond=processed>0?processed/Math.Max(1,watch.Elapsed.TotalSeconds):0;
            if(progress.ItemsPerSecond>0)progress.RemainingSeconds=progress.PendingItems/progress.ItemsPerSecond;
            Console.WriteLine("FLUXO="+JsonSerializer.Serialize(progress));
        }
        workflowProgress=()=>Report("Validando e concluindo as mensagens do lote");
        try
        {
            Report("Preparando a retomada dos lotes",true);
            foreach(var batch in batches)
            {
                // Current item states can change through advanced actions after
                // a batch checkpoint. Reopen only work explicitly requested.
                var batchRows=ledger.WorkflowItems(jobId,batch.Items.ToHashSet(StringComparer.Ordinal));
                bool hasWork=batchRows.Any(x=>WorkflowPlan.Actionable(x.State,move));
                bool hasReview=reconcilePending && batchRows.Any(x=>x.State==ArchiveItemState.ReviewRequired);
                if(!hasWork && !hasReview)continue;
                if(pauseRequested?.Invoke()==true || File.Exists(request))
                {ledger.SetWorkflowState(jobId,batch.Number,"Paused");Report("Pausado antes do próximo lote",true);return 4;}
                current=batch.Number;workflowSelection=batch.Items.ToHashSet(StringComparer.Ordinal);
                var rows=WorkItems(jobId);
                if(hasReview || reconcilePending && rows.Any(x=>x.State is ArchiveItemState.ImportStarted or ArchiveItemState.DeleteStarted))
                {
                    ledger.SetWorkflowState(jobId,batch.Number,"Running");Report("Reconciliando as pendências do lote",true);
                    Reconcile(jobId,workflowSelection);rows=WorkItems(jobId);
                }
                if(!rows.Any(x=>WorkflowPlan.Actionable(x.State,move)))
                {
                    ledger.SetWorkflowState(jobId,batch.Number,WorkflowPlan.Classify(rows,move));
                    Report("Reconciliação concluída; divergências preservadas para revisão",true);
                    if(pauseRequested?.Invoke()==true || File.Exists(request))return 4;
                    continue;
                }
                ledger.SetWorkflowState(jobId,batch.Number,"Running");Report("Copiando as mensagens do lote",true);
                if(rows.Any(x=>x.State is ArchiveItemState.Planned or ArchiveItemState.Exported or ArchiveItemState.ImportStarted))Copy(jobId,moveInline:move,allowRetainedBackup:allowRetainedBackup);
                Report("Validando as cópias e concluindo o lote",true);
                Verify(jobId,allowRetainedBackup:allowRetainedBackup,pendingOnly:true,moveValidatedToDeletedItems:move);
                rows=WorkItems(jobId);string state=WorkflowPlan.Classify(rows,move);
                ledger.SetWorkflowState(jobId,batch.Number,state=="Pending"?"Blocked":state);
                Fault?.Invoke("after-workflow-batch-checkpoint",rows.First());
                Report("Organizando as evidências dos itens concluídos",true);
                MaintainStorage(jobId);
                if(state=="Pending")
                {
                    Report("Execução interrompida por bloqueio; o lote ainda não foi concluído",true);
                    string reason=rows.Any(x=>x.Error=="destination_durable_reopen_or_independent_backup_missing")
                        ? "O Outlook mantém o PST ocupado e a prova adicional com backup MSG retido não está autorizada nesta execução. Feche as mensagens abertas desse PST e tente novamente. Se preferir autorizar a prova adicional, marque Permitir validação com backup MSG retido nesta execução, em Ações avançadas."
                        : string.Join("; ",rows.Where(x=>WorkflowPlan.Actionable(x.State,move)&&x.Error is not null).Select(x=>x.Error).Distinct().Take(3));
                    if(string.IsNullOrWhiteSpace(reason))reason="Há itens do lote que não puderam ser concluídos nesta tentativa. Confira o Outlook e as pendências antes de retomar.";
                    throw new ArchiveReviewException($"Lote {batch.Number} não concluído: {reason} O progresso foi preservado; a execução não está em pausa voluntária.");
                }
                Report(state=="Review"?"Lote encerrado; pendências separadas para revisão":"Lote concluído",true);
                if(pauseRequested?.Invoke()==true || File.Exists(request))
                {
                    var next=batches.FirstOrDefault(x=>x.Number>batch.Number && x.State is not ("Complete" or "Review"));
                    if(next is not null)ledger.SetWorkflowState(jobId,next.Number,"Paused");
                    Report("Pausado após concluir o lote atual",true);return 4;
                }
            }
            var final=ledger.WorkflowStatus(jobId)!;
            Report(final.ReviewItems>0?"Processamento encerrado; confira as pendências agrupadas":"Arquivamento concluído",true);
            return final.ReviewItems>0?3:0;
        }
        finally{workflowSelection=null;workflowProgress=null;}
    }
}

internal sealed partial class OutlookArchiveBackend
{
    internal void CheckWorkflowReady(ArchiveConfig config,IReadOnlyList<ArchiveVolume> volumes)
    {
        object? stores=null;
        try{stores=((dynamic)session!).Stores;_=(int)((dynamic)stores).Count;}finally{Release(stores);}
        foreach(var volume in volumes.Where(x=>x.State==ArchiveVolumeState.Ready))
            if(!File.Exists(volume.PstPath))throw new ArchiveReviewException("O PST registrado não foi encontrado: "+volume.PstPath);
        string destinationFolder=config.DestinationPstPath is { } pst ? Path.GetDirectoryName(pst)! : config.ExistingPstPath is { } existing ? Path.GetDirectoryName(existing)! : config.RootDirectory;
        var folders=new[]{config.RootDirectory,destinationFolder}.Distinct(StringComparer.OrdinalIgnoreCase);
        foreach(string? folder in folders)
        {
            ArchiveFiles.ValidateLocalDirectory(folder!);
            var drive=new DriveInfo(Path.GetPathRoot(Path.GetFullPath(folder!))!);
            if(drive.AvailableFreeSpace<config.ReserveBytes)throw new ArchiveReviewException("Espaço livre abaixo da reserva de segurança em "+drive.Name);
            string probe=Path.Combine(folder!,".archive-access-"+Guid.NewGuid().ToString("N")+".tmp");
            using(var access=new FileStream(probe,FileMode.CreateNew,FileAccess.Write,FileShare.None,1,FileOptions.DeleteOnClose)){access.WriteByte(0);access.Flush(true);}
        }
        Console.WriteLine("Preparação: conexão com Outlook, PSTs e reserva de espaço conferidos.");
    }
}
