using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

internal sealed record PlanningScope(List<FolderScopeSnapshot> Folders,bool Complete,List<FolderScopeSnapshot>? InitialFolders=null);
internal sealed record PlanningEntry(int Ordinal,string Folder,string Entry,SourceDescriptor? Source);

internal sealed partial class ArchiveLedger
{
    private void InitializePlanning()=>Execute("""
        CREATE TABLE IF NOT EXISTS planning_scopes(job TEXT PRIMARY KEY,payload TEXT NOT NULL,sha TEXT NOT NULL,FOREIGN KEY(job) REFERENCES jobs(id));
        CREATE TABLE IF NOT EXISTS planning_entries(job TEXT NOT NULL,ordinal INTEGER NOT NULL,payload TEXT NOT NULL,sha TEXT NOT NULL,PRIMARY KEY(job,ordinal),FOREIGN KEY(job) REFERENCES jobs(id));
        """);
    private static string PlanningSha(string text)=>Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    internal PlanningScope PlanningScope(string job)
    {
        using var cmd=Command("SELECT payload,sha FROM planning_scopes WHERE job=$j",null,("$j",job));using var reader=cmd.ExecuteReader();
        if(!reader.Read())return new([],false);
        string payload=reader.GetString(0);if(PlanningSha(payload)!=reader.GetString(1))throw new ArchiveReviewException("O escopo salvo do planejamento foi alterado.");
        return JsonSerializer.Deserialize<PlanningScope>(payload)??throw new ArchiveReviewException("Escopo salvo inválido.");
    }
    internal void SavePlanningScope(ArchiveJob job,PlanningScope scope,ArchiveConfig config)
    {
        string text=JsonSerializer.Serialize(scope);job.ConfigJson=JsonSerializer.Serialize(new ArchivePlanEnvelope(config,""));
        using var tx=connection.BeginTransaction();
        Run("INSERT INTO planning_scopes(job,payload,sha) VALUES($j,$p,$s) ON CONFLICT(job) DO UPDATE SET payload=excluded.payload,sha=excluded.sha",tx,("$j",job.Id),("$p",text),("$s",PlanningSha(text)));
        Run("UPDATE jobs SET payload=$p WHERE id=$j",tx,("$p",JsonSerializer.Serialize(job)),("$j",job.Id));tx.Commit();
    }
    internal List<PlanningEntry> PlanningEntries(string job)
    {
        var result=new List<PlanningEntry>();
        using var cmd=Command("SELECT ordinal,payload,sha FROM planning_entries WHERE job=$j ORDER BY ordinal",null,("$j",job));using var reader=cmd.ExecuteReader();
        while(reader.Read())
        {
            string text=reader.GetString(1);if(PlanningSha(text)!=reader.GetString(2))throw new ArchiveReviewException("Checkpoint do planejamento foi alterado.");
            var row=JsonSerializer.Deserialize<PlanningEntry>(text)??throw new ArchiveReviewException("Checkpoint inválido.");
            if(row.Ordinal!=reader.GetInt32(0) || row.Ordinal!=result.Count)throw new ArchiveReviewException("Sequência dos checkpoints incompleta.");result.Add(row);
        }
        return result;
    }
    internal int PlanningCount(string job)
    {using var cmd=Command("SELECT COUNT(*) FROM planning_entries WHERE job=$j",null,("$j",job));return Convert.ToInt32(cmd.ExecuteScalar());}
    internal void AddPlanningEntry(string job,PlanningEntry row)
    {
        string text=JsonSerializer.Serialize(row);using var tx=connection.BeginTransaction();
        Run("INSERT INTO planning_entries(job,ordinal,payload,sha) VALUES($j,$o,$p,$s)",tx,("$j",job),("$o",row.Ordinal),("$p",text),("$s",PlanningSha(text)));tx.Commit();
    }
    internal void SavePlanningJob(ArchiveJob job)
    {
        using var tx=connection.BeginTransaction();
        Run("UPDATE jobs SET payload=$p WHERE id=$j",tx,("$p",JsonSerializer.Serialize(job)),("$j",job.Id));
        Event(job.Id,job.PlanningPending?"planning-sealed":"planning-complete",new{job.Id},tx);tx.Commit();
    }
}

internal sealed partial class ArchiveEngine
{
    internal Action<string,int>? PlanningFault;
    private string BeginPlanning(ArchiveConfig config,List<FolderScopeSnapshot>? initialScope=null,IReadOnlyList<FolderScopeSnapshot>? selectedScope=null)
    {
        string id=Guid.NewGuid().ToString("N"),directory=Path.Combine(root,"jobs",id);
        Directory.CreateDirectory(directory);Directory.CreateDirectory(Path.Combine(directory,"psts"));
        var job=new ArchiveJob{Id=id,RootDirectory=directory,PlanningPending=true,ConfigJson=JsonSerializer.Serialize(new ArchivePlanEnvelope(config,""))};
        ledger.InsertJob(job);
        if(initialScope is not null)ledger.SavePlanningScope(job,new(selectedScope!.ToList(),true,initialScope),config);
        Console.WriteLine("PLANEJAMENTO="+id);return ResumePlanning(id);
    }
    internal string ResumePlanning(string id)
    {
        var job=ledger.GetJob(id);
        if(!job.PlanningPending){SealPlan(job);return id;}
        var config=Envelope(job,true).Config;var scope=ledger.PlanningScope(id);
        if(!scope.Complete)
        {
            var saved=scope.Folders.ToList();
            void Checkpoint(FolderScopeSnapshot folder)
            {
                if(saved.Any(x=>x.FolderId==folder.FolderId))throw new ArchiveReviewException("Pasta repetida no planejamento.");
                saved.Add(folder);ledger.SavePlanningScope(job,new(saved,false),config);PlanningFault?.Invoke("scope-checkpoint",saved.Count);
            }
            List<FolderScopeSnapshot> captured;
            if(backend is IArchiveResumableScopeBackend resumable)captured=ArchivePlanningConnection.Read(()=>resumable.CaptureScope(config,saved.ToList(),Checkpoint));
            else
            {
                captured=ArchivePlanningConnection.Read(()=>backend.CaptureScope(config)).Select(folder=>saved.SingleOrDefault(x=>x.FolderId==folder.FolderId)??folder).ToList();
                if(saved.Any(x=>!captured.Any(y=>x.FolderId==y.FolderId)))throw new ArchiveReviewException("Uma pasta capturada não está mais na origem.");
                foreach(var folder in captured)if(!saved.Any(x=>x.FolderId==folder.FolderId))Checkpoint(folder);
            }
            scope=new(captured,true);ledger.SavePlanningScope(job,scope,config);
        }
        var inputs=scope.Folders.SelectMany(f=>f.EntryIds.Select(entry=>(Folder:f,Entry:entry))).ToList();
        if(inputs.Count>config.MaxItems || inputs.Select(x=>ArchiveLedger.ComputeItemKey(config.SourceStoreId,x.Entry)).Distinct(StringComparer.Ordinal).Count()!=inputs.Count)
            throw new ArchiveReviewException("Escopo salvo excede o limite ou contém identidades repetidas.");
        var completed=ledger.PlanningEntries(id);
        if(completed.Count>inputs.Count)throw new ArchiveReviewException("Checkpoints fora do escopo inicial.");
        for(int index=0;index<completed.Count;index++)CheckPlanningEntry(completed[index],inputs[index].Folder,inputs[index].Entry,config);
        Console.WriteLine($"Planejamento: retomando {completed.Count:N0}/{inputs.Count:N0} itens já analisados.");
        for(int index=completed.Count;index<inputs.Count;index++)
        {
            var input=inputs[index];var found=ArchivePlanningConnection.Read(()=>backend.DiscoverFrozen(config,[new(input.Folder.FolderId,input.Folder.RelativePath,[input.Entry])]));
            if(found.Count>1)throw new ArchiveReviewException("Descoberta retornou mais de um item para a identidade congelada.");
            var row=new PlanningEntry(index,input.Folder.FolderId,input.Entry,found.SingleOrDefault());
            CheckPlanningEntry(row,input.Folder,input.Entry,config);ledger.AddPlanningEntry(id,row);completed.Add(row);
            PlanningFault?.Invoke("item-checkpoint",index+1);
            if((index+1)%25==0 || index+1==inputs.Count)Console.WriteLine($"Planejamento: {index+1:N0}/{inputs.Count:N0} itens analisados e salvos.");
        }
        var descriptors=completed.Where(x=>x.Source is not null).Select(x=>x.Source!).ToList();
        void EnsureFile<T>(string path,T value)
        {
            if(!File.Exists(path)){ArchiveFiles.WriteNew(path,value);return;}
            if(JsonSerializer.Serialize(JsonSerializer.Deserialize<T>(File.ReadAllText(path)))!=JsonSerializer.Serialize(value))throw new ArchiveReviewException("Arquivo de planejamento difere dos checkpoints preservados.");
        }
        string plan=Path.Combine(job.RootDirectory,"source-plan.json"),scopePath=Path.Combine(job.RootDirectory,"source-scope.json");
        EnsureFile(scopePath,scope.InitialFolders??scope.Folders);EnsureFile(plan,descriptors);
        job.ConfigJson=JsonSerializer.Serialize(new ArchivePlanEnvelope(config,ArchiveFiles.Sha(plan),ArchiveFiles.Sha(scopePath)));ledger.SavePlanningJob(job);
        PlanningFault?.Invoke("before-catalog",completed.Count);SealPlan(job);
        job.PlanningPending=false;ledger.SavePlanningJob(job);Status(id);return id;
    }
    private static void CheckPlanningEntry(PlanningEntry row,FolderScopeSnapshot folder,string entry,ArchiveConfig config)
    {
        if(row.Folder!=folder.FolderId || row.Entry!=entry)throw new ArchiveReviewException("Checkpoint pertence a outra identidade do escopo.");
        if(row.Source is { } source && (source.StoreId!=config.SourceStoreId || source.FolderId!=folder.FolderId || !string.Equals(source.EntryId,entry,StringComparison.OrdinalIgnoreCase) || !source.RelativePath.SequenceEqual(folder.RelativePath)))
            throw new ArchiveReviewException("Metadados do planejamento não correspondem à identidade congelada.");
    }
}
