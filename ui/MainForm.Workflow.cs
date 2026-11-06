using System.Text.Json;

namespace OutlookArchive.UI;

internal sealed class WorkflowInfo
{
    public string State {get;set;}="Ready";
    public int TotalItems {get;set;}
    public int CompleteItems {get;set;}
    public int ReviewItems {get;set;}
    public int PendingItems {get;set;}
    public int TotalBatches {get;set;}
    public int CompletedBatches {get;set;}
    public int ReviewBatches {get;set;}
    public int CurrentBatch {get;set;}
    public double ItemsPerSecond {get;set;}
    public double? RemainingSeconds {get;set;}
    public string Activity {get;set;}="";
    public bool MoveOriginals {get;set;}
}

internal sealed partial class MainForm
{
    private readonly Button continueArchive=Button("Continuar arquivamento",true);
    private readonly Button pauseAfterBatch=Button("Pausar após este lote");
    private readonly CheckBox advancedActions=new(){Text="Ações avançadas",AutoSize=true};
    private readonly CheckBox retainedMsgCurrent=new(){Text="Permitir validação com backup MSG retido nesta execução",AutoSize=true,AccessibleName="Autorizar backup MSG na execução atual"};
    private readonly ComboBox pendingGroups=new(){Dock=DockStyle.Fill,DropDownStyle=ComboBoxStyle.DropDownList,AccessibleName="Grupos de pendências"};
    private bool workflowRunning,loadingPendingGroups,batchPauseRequested;
    private sealed record PendingGroup(string Key,int Count)
    {public override string ToString()=>Key.Length==0?"Todos os motivos":$"{Count:N0} · {Key}";}
    private async Task ContinueArchive()
    {
        if(busy || displayedJob is not { } job)return;
        if(job.PlanningPending){await ResumePreparation(job);return;}
        var action=WizardRules.NextAction(job.States,autoMoveCurrent.Checked);
        if(action==ContinueAction.None)return;
        if(action==ContinueAction.Remove){await Remove();return;}
        workflowRunning=true;batchPauseRequested=false;
        try{await Execute("archive-workflow-run","Conferindo o Outlook, os PSTs e os lotes pendentes…",reconcilePending:action==ContinueAction.Reconcile);}
        finally{workflowRunning=false;UpdateActions();}
    }
    private async Task ResumePreparation(JobStatus job)
    {
        if(activeRoot is null)return;
        string directory=activeRoot;
        await Guarded("Retomando o planejamento salvo…",async()=>
        {
            preparingJobId=job.Id;
            EnsureSuccess(await engine.RunAsync(["archive-plan-resume",directory,job.Id],AppendLog));
            EnsureSuccess(await engine.RunAsync(["archive-workflow-prepare",directory,job.Id,autoMoveCurrent.Checked?"--move-validated-to-trash":"--keep-originals"],AppendLog));
            await WaitForStatusPoll();await RefreshJobs(job.Id);ShowStep(5);
            activity.Text="Planejamento concluído. Use Continuar arquivamento para iniciar os lotes.";
        });
    }
    private void RequestBatchPause()
    {
        if(!busy || !workflowRunning || displayedJob is not { } job || activeRoot is null)return;
        try
        {
            string folder=Path.Combine(activeRoot,"jobs",job.Id,"workflow");Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder,"pause-after-batch.request"),DateTimeOffset.UtcNow.ToString("O"));
            batchPauseRequested=true;activity.Text="Pausa solicitada. O lote atual será concluído antes de parar.";pauseAfterBatch.Enabled=false;
        }
        catch(Exception ex) when(ex is IOException or UnauthorizedAccessException){MessageBox.Show(this,ex.Message,"Não foi possível solicitar pausa");}
    }
    private void ShowWorkflowSummary(JobStatus job)
    {
        int copied=job.States.Where(x=>x.Key is "Copied" or "Verified" or "InDeletedItems" or "Removed" or "BackupVerifiedSourceUnmanaged").Sum(x=>x.Value);
        int reviews=job.Workflow?.ReviewItems??job.Errors.Count;
        string text=$"Origem: {job.SourceLabel}\n\nCópias registradas: {copied:N0}\nOriginais em Itens Excluídos: {job.States.GetValueOrDefault("InDeletedItems"):N0}\nValidados, mantidos na origem: {job.States.GetValueOrDefault("Verified"):N0}\nCópias validadas sem ação na origem: {job.States.GetValueOrDefault("BackupVerifiedSourceUnmanaged"):N0}\nConvites de reunião: {job.States.GetValueOrDefault("SkippedByPolicy"):N0}\nPendências para revisão: {reviews:N0}\n\nLotes concluídos: {job.Workflow?.CompletedBatches}/{job.Workflow?.TotalBatches}.\nConsulte a aba Pendências para filtrar os motivos restantes.";
        string directory=Path.Combine(root.Text,"reports");Directory.CreateDirectory(directory);
        File.WriteAllText(Path.Combine(directory,$"resumo-{job.Id[..8]}-{DateTime.Now:yyyyMMdd-HHmmss}.txt"),text);
        MessageBox.Show(this,text,reviews>0?"Arquivamento encerrado com pendências":"Arquivamento concluído",MessageBoxButtons.OK,MessageBoxIcon.Information);
    }
    private void ApplyWorkflowProgress(WorkflowInfo info,bool activityUpdate)
    {
        progress.Style=ProgressBarStyle.Continuous;progress.Maximum=Math.Max(1,info.TotalItems);progress.Value=Math.Clamp(info.CompleteItems,0,progress.Maximum);
        if(activityUpdate)
        {
            string estimate=info.RemainingSeconds is { } seconds && double.IsFinite(seconds)?$" · estimativa restante: {TimeSpan.FromSeconds(Math.Min(seconds,8640000)):d\\d\\ hh\\h\\ mm\\m}":"";
            string speed=info.ItemsPerSecond>0?$" · {info.ItemsPerSecond:N1} itens/s":"";
            activity.Text=$"Lote {info.CurrentBatch}/{info.TotalBatches} · {info.Activity}\n{info.CompleteItems:N0}/{info.TotalItems:N0} concluídos · {info.ReviewItems:N0} para revisão{speed}{estimate}";
            if(batchPauseRequested)activity.Text+="\nPausa solicitada para o final deste lote.";
            UpdateActions();
        }
    }
    private static string PendingGroupFor(string reason)
    {
        if(reason.Contains("source_changed",StringComparison.OrdinalIgnoreCase)||reason.Contains("target_changed",StringComparison.OrdinalIgnoreCase)||reason.Contains("fields_differ",StringComparison.OrdinalIgnoreCase)||reason.Contains("diverg",StringComparison.OrdinalIgnoreCase)||reason.Contains("não correspond",StringComparison.OrdinalIgnoreCase))return "Diferenças de conteúdo ou metadados";
        if(reason.Contains("não localiz",StringComparison.OrdinalIgnoreCase)||reason.Contains("ausente",StringComparison.OrdinalIgnoreCase))return "Origem ou arquivo não encontrado";
        if(reason.Contains("durable",StringComparison.OrdinalIgnoreCase)||reason.Contains("reopen",StringComparison.OrdinalIgnoreCase))return "Persistência ou disponibilidade do PST";
        if(reason.Contains("suportad",StringComparison.OrdinalIgnoreCase)||reason.Contains("rascunho",StringComparison.OrdinalIgnoreCase))return "Formato de mensagem pendente";
        return reason.Split(':')[0] is {Length:>0} prefix ? prefix[..Math.Min(prefix.Length,100)] : "Outras pendências";
    }
    private void RefreshPendingGroups()
    {
        string selected=(pendingGroups.SelectedItem as PendingGroup)?.Key??"";loadingPendingGroups=true;
        try
        {
            var errors=displayedJob?.Errors.Where(UiRules.IsActionableError).ToList()??[];
            pendingGroups.Items.Clear();pendingGroups.Items.Add(new PendingGroup("",errors.Count));
            pendingGroups.Items.AddRange(errors.GroupBy(x=>PendingGroupFor(x.Error)).OrderByDescending(x=>x.Count()).Select(x=>(object)new PendingGroup(x.Key,x.Count())).ToArray());
            pendingGroups.SelectedItem=pendingGroups.Items.Cast<PendingGroup>().FirstOrDefault(x=>x.Key==selected)??pendingGroups.Items[0];
        }
        finally{loadingPendingGroups=false;}
    }
    private List<JobError> FilteredPending()
    {
        string key=(pendingGroups.SelectedItem as PendingGroup)?.Key??"";
        return displayedJob?.Errors.Where(UiRules.IsActionableError).Where(x=>key.Length==0||PendingGroupFor(x.Error)==key).ToList()??[];
    }
}
