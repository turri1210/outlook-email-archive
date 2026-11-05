using System.Text.Json;
using OpenMcdf;

internal sealed partial class OutlookArchiveBackend
{
    internal void InspectRemainingTypes(string rootDirectory,string jobId,string output)
    {
        using var lease=new ArchiveEngine(rootDirectory,this);
        using var ledger=new ArchiveLedger(Path.Combine(rootDirectory,"archive.sqlite"),true);
        var job=ledger.GetJob(jobId);var results=new List<object>();
        foreach(var row in ledger.GetItems(jobId).Where(x=>x.State==ArchiveItemState.SkippedUnsupported || (x.Error??"").Contains("MSG sem stream de corpo reconhecido",StringComparison.Ordinal)))
        {
            object? item=null,parent=null,attachments=null;
            try
            {
                item=((dynamic)session!).GetItemFromID(row.SourceEntryId,row.SourceStoreId);parent=((dynamic)item).Parent;
                if((string)((dynamic)parent).StoreID!=row.SourceStoreId || (string)((dynamic)parent).EntryID!=row.SourceFolderId || row.SourceRecordKey is not null && BinaryProperty(item,"0x0FF90102")!=row.SourceRecordKey)throw new ArchiveReviewException("Item fora da identidade inicial.");
                int kind=(int)((dynamic)item).Class;string cls=(string)((dynamic)item).MessageClass;
                string folder=Path.Combine(job.RootDirectory,"type-analysis",row.Key[..16]+"-"+Guid.NewGuid().ToString("N")[..8]);Directory.CreateDirectory(folder);
                string file=Path.Combine(folder,"message.msg");((dynamic)item).SaveAs(file,9);
                using var cfb=RootStorage.OpenRead(file);
                var streams=cfb.EnumerateEntries().Select(x=>x.Name).ToArray();
                attachments=((dynamic)item).Attachments;
                results.Add(new{Key=row.Key,Class=kind,MessageClass=cls,Sent=kind==43?(bool?)((dynamic)item).Sent:null,DownloadState=kind==43?(int?)((dynamic)item).DownloadState:null,
                    BodyFormat=kind==43?(int?)((dynamic)item).BodyFormat:null,BodyLength=kind==43?((string?)((dynamic)item).Body)?.Length:null,
                    HtmlLength=kind==43?((string?)((dynamic)item).HTMLBody)?.Length:null,Attachments=(int)((dynamic)attachments).Count,Streams=streams,Message=file,MessageSha=ArchiveFiles.Sha(file),OriginalTouched=false});
            }
            catch(Exception ex){results.Add(new{Key=row.Key,Error=ex.Message});}
            finally{Release(attachments);Release(parent);Release(item);}
        }
        ArchiveFiles.WriteNew(Path.GetFullPath(output),results);Console.WriteLine("Tipos e corpos restantes inspecionados sem alterar os originais.");
    }
}
