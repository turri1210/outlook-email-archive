using System.Text.Json;

internal sealed record WorkflowBatch(int Number,List<string> Items,string State);
internal sealed class WorkflowProgress
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

internal sealed partial class ArchiveLedger
{
    private void InitializeWorkflow()=>Execute("""
        CREATE TABLE IF NOT EXISTS workflow_runs(job TEXT PRIMARY KEY,manifest TEXT NOT NULL,sha TEXT NOT NULL,move INTEGER NOT NULL,FOREIGN KEY(job) REFERENCES jobs(id));
        CREATE TABLE IF NOT EXISTS workflow_batches(job TEXT NOT NULL,number INTEGER NOT NULL,state TEXT NOT NULL,items TEXT NOT NULL,PRIMARY KEY(job,number),FOREIGN KEY(job) REFERENCES jobs(id));
        CREATE TABLE IF NOT EXISTS workflow_members(key TEXT PRIMARY KEY,job TEXT NOT NULL,batch INTEGER NOT NULL,FOREIGN KEY(key) REFERENCES items(key),FOREIGN KEY(job,batch) REFERENCES workflow_batches(job,number));
        CREATE INDEX IF NOT EXISTS workflow_state ON workflow_batches(job,state,number);
        """);
    private bool HasWorkflowTable()
    {
        using var cmd=Command("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='workflow_runs'",null);
        return Convert.ToInt64(cmd.ExecuteScalar())>0;
    }
    internal (string Path,string Sha,bool Move)? WorkflowManifest(string job)
    {
        if(!HasWorkflowTable())return null;
        using var cmd=Command("SELECT manifest,sha,move FROM workflow_runs WHERE job=$job",null,("$job",job));using var reader=cmd.ExecuteReader();
        return reader.Read() ? (reader.GetString(0),reader.GetString(1),reader.GetInt64(2)!=0) : null;
    }
    internal List<WorkflowBatch> WorkflowBatches(string job)
    {
        if(WorkflowManifest(job) is null)return [];
        using var cmd=Command("SELECT number,items,state FROM workflow_batches WHERE job=$job ORDER BY number",null,("$job",job));using var reader=cmd.ExecuteReader();var result=new List<WorkflowBatch>();
        while(reader.Read())result.Add(new(reader.GetInt32(0),JsonSerializer.Deserialize<List<string>>(reader.GetString(1))!,reader.GetString(2)));
        return result;
    }
    internal List<ArchiveItem> WorkflowItems(string job,IReadOnlySet<string> keys)
    {
        if(keys.Count==0)return [];
        var values=keys.Select((key,index)=>("$k"+index,(object?)key)).ToList();values.Add(("$job",job));
        return Many<ArchiveItem>("SELECT payload FROM items WHERE job=$job AND key IN ("+string.Join(",",values.Take(keys.Count).Select(x=>x.Item1))+") ORDER BY rowid",values.ToArray());
    }
    internal void AddWorkflow(string job,string manifest,string sha,bool move,List<WorkflowBatch> batches)
    {
        using var tx=connection.BeginTransaction();
        using(var version=Command("PRAGMA application_id",tx))
            if(Convert.ToInt64(version.ExecuteScalar())!=1128355893)Run("PRAGMA application_id=1128355892; PRAGMA user_version=4;",tx);
        Run("INSERT INTO workflow_runs(job,manifest,sha,move) VALUES($j,$p,$s,$m)",tx,("$j",job),("$p",manifest),("$s",sha),("$m",move?1:0));
        foreach(var batch in batches)
        {
            Run("INSERT INTO workflow_batches(job,number,state,items) VALUES($j,$n,$s,$i)",tx,("$j",job),("$n",batch.Number),("$s",batch.State),("$i",JsonSerializer.Serialize(batch.Items)));
            foreach(string key in batch.Items)Run("INSERT INTO workflow_members(key,job,batch) VALUES($k,$j,$n)",tx,("$k",key),("$j",job),("$n",batch.Number));
        }
        Event(job,"workflow-prepared",new{manifest,sha,move,Batches=batches.Count},tx);tx.Commit();
    }
    internal int RequeueUnavailableNativeIdentity(string job)
    {
        var rows=GetItems(job).Where(x=>x.State==ArchiveItemState.ReviewRequired && x.Error=="Identificação nativa da mensagem pendente: message unavailable" &&
            x.StageMsgPath is not null && x.StageMsgSha256 is not null && x.SnapshotHash is not null && x.SourceRelevantHash is not null &&
            x.TargetEntryId is not null && x.TargetSnapshotHash is not null && x.TargetFieldsPath is not null && x.TargetRelevantHash is not null).ToList();
        if(rows.Count==0)return 0;
        using var tx=connection.BeginTransaction();
        foreach(var row in rows)
        {
            string reason=row.Error!;int previous=row.Version;row.State=ArchiveItemState.Copied;row.Error=null;row.Version++;
            if(Run("UPDATE items SET state=$s,version=$v,payload=$p WHERE key=$k AND version=$old",tx,("$s",(int)row.State),("$v",row.Version),("$p",JsonSerializer.Serialize(row)),("$k",row.Key),("$old",previous))!=1)
                throw new InvalidOperationException("Conflito ao reenfileirar cópia para validação.");
            Run("UPDATE workflow_batches SET state='Pending' WHERE job=$j AND number IN (SELECT batch FROM workflow_members WHERE key=$k)",tx,("$j",job),("$k",row.Key));
            Event(job,"workflow-revalidation-queued",new{row.Key,PreviousError=reason,NewState="Copied",ApprovalGranted=false},tx);
        }
        tx.Commit();return rows.Count;
    }
    internal void SetWorkflowState(string job,int number,string state)
    {
        if(state is not ("Pending" or "Running" or "Paused" or "Blocked" or "Complete" or "Review"))throw new ArgumentException("Estado de lote inválido.");
        using var tx=connection.BeginTransaction();Run("UPDATE workflow_batches SET state=$s WHERE job=$j AND number=$n",tx,("$s",state),("$j",job),("$n",number));
        Event(job,"workflow-batch-state",new{Number=number,State=state},tx);tx.Commit();
    }
    internal void SetWorkflowMovement(string job,bool move)
    {
        if(WorkflowManifest(job)?.Move==move)return;
        using var tx=connection.BeginTransaction();
        // Changing the goal to moving originals reopens only batches that
        // still contain validated originals, never completed trash movements.
        if(move)Run("UPDATE workflow_batches SET state='Pending' WHERE job=$j AND state='Complete' AND number IN (SELECT w.batch FROM workflow_members w JOIN items i ON i.key=w.key WHERE w.job=$j AND i.state=4)",tx,("$j",job));
        Run("UPDATE workflow_runs SET move=$m WHERE job=$j",tx,("$m",move?1:0),("$j",job));Event(job,"workflow-goal-changed",new{MoveOriginals=move},tx);tx.Commit();
    }
    internal WorkflowProgress? WorkflowStatus(string job)
    {
        var manifest=WorkflowManifest(job);if(manifest is null)return null;
        using var cmd=Command("SELECT state,COUNT(*) FROM items WHERE job=$j GROUP BY state",null,("$j",job));using var reader=cmd.ExecuteReader();var counts=new Dictionary<ArchiveItemState,int>();
        while(reader.Read())counts[(ArchiveItemState)reader.GetInt32(0)]=reader.GetInt32(1);
        var batches=WorkflowBatches(job);
        int complete=counts.Where(x=>WorkflowPlan.Done(x.Key,manifest.Value.Move)).Sum(x=>x.Value);
        int reviews=counts.Where(x=>!WorkflowPlan.Done(x.Key,manifest.Value.Move)&&!WorkflowPlan.Actionable(x.Key,manifest.Value.Move)).Sum(x=>x.Value);
        var current=batches.FirstOrDefault(x=>x.State is "Running" or "Paused" or "Blocked");
        return new(){MoveOriginals=manifest.Value.Move,TotalItems=counts.Values.Sum(),CompleteItems=complete,ReviewItems=reviews,PendingItems=counts.Values.Sum()-complete-reviews,TotalBatches=batches.Count,CompletedBatches=batches.Count(x=>x.State=="Complete"),ReviewBatches=batches.Count(x=>x.State=="Review"),CurrentBatch=current?.Number??0,
            State=current?.State ?? (batches.Any(x=>x.State=="Pending")?"Ready":reviews>0?"Review":"Complete")};
    }
}
