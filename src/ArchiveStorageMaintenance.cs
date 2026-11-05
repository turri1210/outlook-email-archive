using System.Text.Json;

internal sealed partial class ArchiveLedger
{
    private void InitializeEvidenceMaintenance()=>Execute("CREATE TABLE IF NOT EXISTS evidence_maintenance(key TEXT PRIMARY KEY,version INTEGER NOT NULL,utc TEXT NOT NULL,FOREIGN KEY(key) REFERENCES items(key));");
    internal bool EvidenceNeedsMaintenance(ArchiveItem item)
    {
        using var command=Command("SELECT version FROM evidence_maintenance WHERE key=$k",null,("$k",item.Key));var value=command.ExecuteScalar();return value is null||Convert.ToInt32(value)!=item.Version;
    }
    internal void MarkEvidenceMaintenance(ArchiveItem item)=>Run("INSERT INTO evidence_maintenance(key,version,utc) VALUES($k,$v,$u) ON CONFLICT(key) DO UPDATE SET version=excluded.version,utc=excluded.utc",null,("$k",item.Key),("$v",item.Version),("$u",DateTimeOffset.UtcNow.ToString("O")));
}

internal sealed partial class ArchiveEngine
{
    internal object MaintainStorage(string jobId,bool reclaim=false)
    {
        var job=ledger.GetJob(jobId);bool move=AutomaticMovementEnabled(jobId);
        return ArchiveStorageMaintenance.Run(ledger,job,WorkItems(jobId),move,reclaim);
    }
}

internal static class ArchiveStorageMaintenance
{
    internal static object Run(ArchiveLedger ledger,ArchiveJob job,IEnumerable<ArchiveItem> items,bool move,bool reclaim)
    {
        var rows=items.Where(x=>WorkflowPlan.Done(x.State,move)&&ledger.EvidenceNeedsMaintenance(x)).ToList();
        int compressed=0,removed=0,skipped=0;long saved=0;
        foreach(var row in rows)
        {
            var result=ArchiveEvidenceStorage.Organize(job,[row],move);compressed+=result.FilesCompressed;removed+=result.TemporaryFilesRemoved;skipped+=result.BusyOrUnsupported;saved+=result.BytesSaved;
            if(result.BusyOrUnsupported==0)ledger.MarkEvidenceMaintenance(row);
        }
        int packed=reclaim?ledger.PackHistory(job.Id):0;
        if(!reclaim)ledger.MaybePackHistory(job.Id);
        if(reclaim)ledger.ReclaimHistorySpace();
        var summary=new{JobId=job.Id,FilesCompressed=compressed,TemporaryFilesRemoved=removed,BusyOrUnsupported=skipped,EvidenceBytesSaved=saved,HistoryEventsPacked=packed,Reclaimed=reclaim,Utc=DateTimeOffset.UtcNow};
        string directory=Path.Combine(job.RootDirectory,"storage");Directory.CreateDirectory(directory);ArchiveFiles.WriteNew(Path.Combine(directory,"maintenance-"+Guid.NewGuid().ToString("N")+".json"),summary);
        Console.WriteLine("Armazenamento: "+compressed+" arquivos comprimidos; "+removed+" cópias temporárias removidas; "+(saved/1024d/1024).ToString("N1")+" MiB recuperados nas evidências.");return summary;
    }
}
