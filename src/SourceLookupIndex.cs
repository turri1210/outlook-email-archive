internal sealed record SourceIndexRow(string EntryId,string? RecordKey,string? Token);

internal sealed class SourceLookupIndex
{
    private readonly Dictionary<string,SourceIndexRow> entries=new(StringComparer.Ordinal);
    private readonly Dictionary<string,HashSet<string>> records=new(StringComparer.Ordinal),tokens=new(StringComparer.Ordinal);
    internal int Count=>entries.Count;
    private static string Key(string value)=>value.Length%2==0&&value.All(Uri.IsHexDigit)?value.ToUpperInvariant():value;
    internal void Add(SourceIndexRow row)
    {
        string id=Key(row.EntryId);Remove(id);entries[id]=row;
        void Link(Dictionary<string,HashSet<string>> map,string value){if(!map.TryGetValue(value,out var set))map[value]=set=[];set.Add(id);}
        if(row.RecordKey is not null)Link(records,Key(row.RecordKey));if(row.Token is not null)Link(tokens,row.Token);
    }
    internal void Remove(string entry)
    {
        string id=Key(entry);if(!entries.Remove(id,out var row))return;
        void Unlink(Dictionary<string,HashSet<string>> map,string value){if(map.TryGetValue(value,out var set)){set.Remove(id);if(set.Count==0)map.Remove(value);}}
        if(row.RecordKey is not null)Unlink(records,Key(row.RecordKey));if(row.Token is not null)Unlink(tokens,row.Token);
    }
    internal SourceIndexRow? ByRecord(string key)=>Unique(records,Key(key));
    internal SourceIndexRow? ByToken(string token)=>Unique(tokens,token);
    private SourceIndexRow? Unique(Dictionary<string,HashSet<string>> map,string key)
    {if(!map.TryGetValue(key,out var set))return null;if(set.Count>1)throw new ArchiveReviewException("Identificação ambígua no índice da origem; movimento bloqueado.");return entries[set.Single()];}
}

internal sealed partial class OutlookArchiveBackend
{
    private sealed record FolderIndex(SourceLookupIndex Index,int Generation);
    private readonly Dictionary<string,FolderIndex> sourceLookupIndexes=new(StringComparer.OrdinalIgnoreCase);
    private int lookupGeneration;
    internal void BeginSourceLookupBatch()=>lookupGeneration++;
    private static string LookupKey(string store,string folder)=>store+"|"+folder;
    private FolderIndex ReadSourceIndex(object folder)
    {
        object? table=null,columns=null;
        try
        {
            string store=(string)((dynamic)folder).StoreID,id=(string)((dynamic)folder).EntryID;
            table=((dynamic)folder).GetTable("",0);columns=((dynamic)table).Columns;
            ((dynamic)columns).RemoveAll();((dynamic)columns).Add("EntryID");((dynamic)columns).Add("http://schemas.microsoft.com/mapi/proptag/0x0FF90102");
            bool marker=true;
            try{((dynamic)columns).Add(RemovalTokenProperty+"/0x0000001f");}
            catch(System.Runtime.InteropServices.COMException ex) when(ex.HResult==unchecked((int)0x8004010F)||ex.HResult==unchecked((int)0x80020009)){marker=false;}
            var index=new SourceLookupIndex();int scanned=0;
            while(!(bool)((dynamic)table).EndOfTable)
            {
                Array rows=(Array)((dynamic)table).GetArray(250);
                for(int i=rows.GetLowerBound(0);i<=rows.GetUpperBound(0);i++)
                {
                    int start=rows.GetLowerBound(1);string entry=rows.GetValue(i,start) as string??throw new ArchiveReviewException("Linha de índice sem identificador.");
                    string? key=rows.GetValue(i,start+1) is byte[] bytes?Convert.ToHexString(bytes):null;
                    string? token=marker?rows.GetValue(i,start+2) as string:null;
                    index.Add(new(entry,key,token));scanned++;
                }
                if(scanned>1_000_000)throw new ArchiveReviewException("Índice da origem excede o limite de itens.");
            }
            if(index.Count!=scanned)throw new ArchiveReviewException("Índice de origem contém identificadores repetidos.");
            var result=new FolderIndex(index,lookupGeneration);sourceLookupIndexes[LookupKey(store,id)]=result;
            Console.WriteLine($"Localização: índice de {scanned:N0} mensagens carregado em lote.");return result;
        }
        finally{Release(columns);Release(table);}
    }
    private object? FindIndexed(object folder,string value,bool token=false)
    {
        string store=(string)((dynamic)folder).StoreID,id=(string)((dynamic)folder).EntryID;
        string cacheKey=LookupKey(store,id);
        if(!sourceLookupIndexes.TryGetValue(cacheKey,out var cache))cache=ReadSourceIndex(folder);
        SourceIndexRow? Find()=>token?cache.Index.ByToken(value):cache.Index.ByRecord(value);
        var row=Find();
        // At most one negative refresh per folder and batch. Never infer a
        // missing original from a stale process-wide negative result.
        if(row is null&&cache.Generation!=lookupGeneration){cache=ReadSourceIndex(folder);row=Find();}
        if(row is null)return null;
        var item=TryItem(row.EntryId,store);
        if(item is not null)return item;
        cache=ReadSourceIndex(folder);row=Find();return row is null?null:TryItem(row.EntryId,store);
    }
    private void UpdateSourceIndex(ArchiveItem item,SourceLocation moved,string oldEntry)
    {
        if(sourceLookupIndexes.TryGetValue(LookupKey(item.SourceStoreId,item.SourceFolderId),out var source))source.Index.Remove(oldEntry);
        if(sourceLookupIndexes.TryGetValue(LookupKey(item.SourceStoreId,moved.FolderId),out var trash))trash.Index.Add(new(moved.EntryId,moved.RecordKey,item.Token));
    }
}
