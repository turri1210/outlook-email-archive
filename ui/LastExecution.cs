using System.Text.Json;

namespace OutlookArchive.UI;
internal static class LastExecution
{
    private static string FileFor(string root)=>Path.Combine(root,"preferences","last-execution.json");
    internal static string? Read(string root)
    {
        try
        {
            string file=FileFor(root);if(!File.Exists(file))return null;
            string? id=JsonSerializer.Deserialize<string>(File.ReadAllText(file));return Guid.TryParseExact(id,"N",out _)?id:null;
        }
        catch(Exception ex) when(ex is IOException or UnauthorizedAccessException or JsonException){return null;}
    }
    internal static void Save(string root,string id)
    {
        if(!Guid.TryParseExact(id,"N",out _))return;
        string file=FileFor(root),temp=file+"."+Guid.NewGuid().ToString("N")+".tmp";
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);
            File.WriteAllText(temp,JsonSerializer.Serialize(id));File.Move(temp,file,true);
        }
        catch(Exception ex) when(ex is IOException or UnauthorizedAccessException){/* Optional UI preference must never interrupt a worker. */}
        finally{try{if(File.Exists(temp))File.Delete(temp);}catch(IOException){}catch(UnauthorizedAccessException){}}
    }
    internal static JobStatus? Select(IEnumerable<JobStatus> jobs,string? explicitId,string? currentId,string? rememberedId)
    {
        var rows=jobs.ToList();
        return rows.FirstOrDefault(x=>x.Id==explicitId) ?? rows.FirstOrDefault(x=>x.Id==currentId) ?? rows.FirstOrDefault(x=>x.Id==rememberedId)
            ?? rows.FirstOrDefault(x=>x.PlanningPending||WizardRules.CanContinue(x.States)) ?? rows.FirstOrDefault();
    }
}
