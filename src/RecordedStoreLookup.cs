using System.Runtime.InteropServices;

// A recorded destination is resolved directly. Never enumerate unrelated PSTs
// to find a store whose identity is already part of the durable catalog.
internal static class RecordedStoreLookup
{
    internal static object? Open(string expectedPath, string storeId, Func<string, object?> open,
        Func<object, string> pathOf, Action<object?> release)
    {
        object? store = null;
        try
        {
            store = open(storeId);
            if (store is null) return null;
            string path = pathOf(store);
            if (string.IsNullOrWhiteSpace(path) || !string.Equals(Path.GetFullPath(expectedPath), Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase))
                throw new ArchiveReviewException("O identificador registrado não corresponde ao PST previsto; outros arquivos não serão abertos.");
            object result = store; store = null; return result;
        }
        // Outlook's late-bound automation can wrap MAPI not-found in
        // DISP_E_EXCEPTION. A failed lookup grants no identity; the caller
        // must resolve the exact PST path and validate it again.
        catch (COMException ex) when (ex.HResult is unchecked((int)0x8004010F) or unchecked((int)0x80040107) or unchecked((int)0x80020009)) { return null; }
        finally { release(store); }
    }
    internal static int SelfTest()
    {
        var opened = new List<string>(); var target = new object(); int released = 0;
        var result = Open(@"C:\Archive\real.pst", "recorded-target", id => { opened.Add(id); if(id!="recorded-target")throw new Exception("Unrelated store opened"); return target; }, _ => @"c:\archive\REAL.PST", x => { if(x is not null)released++; });
        if(!ReferenceEquals(result,target)||opened.Count!=1||released!=0)throw new Exception("Recorded PST lookup failed");
        if(Open(@"C:\Archive\real.pst","missing",_=>throw new COMException("Missing",unchecked((int)0x8004010F)),_=>throw new Exception(),_=>{}) is not null)throw new Exception("Missing store accepted");
        if(Open(@"C:\Archive\real.pst","wrapped-missing",_=>throw new COMException("Missing",unchecked((int)0x80020009)),_=>throw new Exception(),_=>{}) is not null)throw new Exception("Wrapped missing store accepted");
        bool rejected=false;
        try{Open(@"C:\Archive\real.pst","wrong",_=>target,_=>@"C:\Other\test.pst",x=>{if(x is not null)released++;});}catch(ArchiveReviewException){rejected=true;}
        if(!rejected||released!=1)throw new Exception("Unrelated PST path accepted or handle leaked");
        Console.WriteLine("PST registrado: acesso direto, ausência e destino diferente conferidos sem enumerar outros arquivos.");return 0;
    }
}
