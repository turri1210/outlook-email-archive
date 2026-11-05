using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

// Isolated synthetic experiment, deliberately not reachable from archive-run.
internal static class NativeBatchProbe
{
    internal static int Execute(string[] args)
    {
        if (args.Length is < 2 or > 3) { Console.Error.WriteLine("Uso: native-batch-probe <manifest.json> <novo-diretório> [lotes-1-a-20-ou-lote-100]"); return 2; }
        object? app=null, session=null, sourceStore=null, source=null, targetStore=null, root=null;
        string? run=null; var batches=new List<object>(); string? error=null; bool allMatched=true,runCreated=false;
        var total=Stopwatch.StartNew();
        try
        {
            int requested=args.Length==3 ? int.Parse(args[2]) : 1;
            bool largeBatch=requested==100; int repeats=largeBatch?1:requested;
            if(repeats is < 1 or > 20) throw new ArgumentException("Limite de vinte lotes artificiais.");
            string manifestPath=Path.GetFullPath(args[0]);
            var manifest=JsonSerializer.Deserialize<Manifest>(File.ReadAllText(manifestPath)) ?? throw new InvalidDataException("Manifesto inválido.");
            string fixtureRoot=Path.GetDirectoryName(manifestPath)!;
            string expectedSource=Path.Combine(fixtureRoot,"source-fixtures.pst");
            if(!manifest.SyntheticOnly || manifest.SendsEmail || manifest.RemovalAllowed || manifest.Recipient!="fixture@example.invalid" ||
                manifest.Items.Count!=5 || !Path.GetFullPath(manifest.SourcePstPath).Equals(expectedSource,StringComparison.OrdinalIgnoreCase) || !File.Exists(expectedSource))
                throw new InvalidDataException("Somente o conjunto artificial original de cinco mensagens é permitido.");
            run=Path.GetFullPath(args[1]); string destination=Path.Combine(run,"native-copy-only.pst");
            if(Directory.Exists(run)||File.Exists(run)) throw new IOException("Use diretório novo.");
            PilotCommand.ValidateDestination(Path.Combine(Path.GetDirectoryName(run)!,"native-probe-validation-"+Guid.NewGuid().ToString("N")+".pst"),run,5);
            string helper=Path.Combine(AppContext.BaseDirectory,"native-copy.exe");
            if(!File.Exists(helper))throw new FileNotFoundException("Helper C++ isolado ausente.",helper);
            if(!CreateDirectoryW(run,IntPtr.Zero))throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(),"Diretório experimental novo não pôde ser criado.");
            runCreated=true;
            var clsid=new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid,IntPtr.Zero,out app)); session=((dynamic)app!).Session;
            sourceStore=FindStore(session!,expectedSource);
            if(sourceStore is null) { ((dynamic)session!).AddStoreEx(expectedSource,2); sourceStore=FindStore(session!,expectedSource); }
            if(sourceStore is null || !(bool)((dynamic)sourceStore).IsDataFileStore ||
                (string)((dynamic)sourceStore).StoreID!=manifest.SourceStoreId)throw new InvalidDataException("PST artificial não corresponde ao manifesto.");
            source=((dynamic)session!).GetFolderFromID(manifest.SourceFolderId,manifest.SourceStoreId);
            GuardFolder(source!,manifest.SourceStoreId,manifest.SourceFolderId);
            var before=new Dictionary<string,string>(); var originalFields=new Dictionary<string,Dictionary<string,string>>();
            long sourceBytes=0; int sourceCount=Count(source!);
            var snapshotTimer=Stopwatch.StartNew();
            foreach(var entry in manifest.Items)
            {
                object? item=null;
                try
                {
                    item=((dynamic)session!).GetItemFromID(entry.EntryId,manifest.SourceStoreId);
                    GuardItem(item!,manifest.SourceStoreId,manifest.SourceFolderId);
                    string subject=(string)((dynamic)item).Subject;
                    if(subject!=$"Teste local Unicode {entry.Index:D2} — ação e café" || (string)((dynamic)item).To!="fixture@example.invalid")throw new InvalidDataException("Item não artificial.");
                    string directory=Path.Combine(run,$"source-before-{entry.Index:D2}");
                    before[entry.EntryId]=Snapshot(item!,directory); originalFields[subject]=ReadFields(directory);
                    sourceBytes+=(int)((dynamic)item).Size;
                }
                finally { Release(item); }
            }
            double beforeSeconds=snapshotTimer.Elapsed.TotalSeconds;
            if(File.Exists(destination))throw new IOException("Destino passou a existir.");
            ((dynamic)session!).AddStoreEx(destination,2); targetStore=FindStore(session!,destination) ?? throw new IOException("Destino não encontrado.");
            root=((dynamic)targetStore).GetRootFolder(); string targetStoreId=(string)((dynamic)targetStore).StoreID;
            for(int batch=1;batch<=repeats;batch++)
            {
                object? folders=null, target=null; var copied=new List<object>();
                try
                {
                    folders=((dynamic)root).Folders; target=((dynamic)folders).Add($"Lote artificial {batch:D2}",6);
                    string targetId=(string)((dynamic)target).EntryID;
                    if(Count(target!)!=0)throw new InvalidDataException("Pasta de teste não vazia.");
                    var handoff=new[]{"synthetic-copy-only-v1",manifest.SourceStoreId,manifest.SourceFolderId,targetStoreId,targetId}.Concat(manifest.Items.Select(x=>x.EntryId));
                    string handoffPath=Path.Combine(run,$"batch-{batch:D2}.txt"); File.WriteAllLines(handoffPath,handoff,new UTF8Encoding(false));
                    var processTimer=Stopwatch.StartNew();
                    using var process=new Process{StartInfo=new ProcessStartInfo(helper){UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true}};
                    process.StartInfo.ArgumentList.Add(handoffPath); process.Start();
                    string stdout=process.StandardOutput.ReadToEnd(); string stderr=process.StandardError.ReadToEnd(); process.WaitForExit(); processTimer.Stop();
                    File.WriteAllText(Path.Combine(run,$"batch-{batch:D2}-helper.json"),stdout);
                    File.WriteAllText(Path.Combine(run,$"batch-{batch:D2}-helper-errors.txt"),stderr);
                    Console.WriteLine($"Lote {batch}: {stdout.Trim()}");
                    var verification=Stopwatch.StartNew(); object? items=null;
                    try
                    {
                        items=((dynamic)target).Items;
                        for(int i=1;i<=(int)((dynamic)items).Count;i++)
                        {
                            object? item=null;
                            try
                            {
                                item=((dynamic)items)[i]; GuardItem(item!,targetStoreId,targetId);
                                string subject=(string)((dynamic)item).Subject; string dir=Path.Combine(run,$"target-{batch:D2}-{i:D2}");
                                Snapshot(item!,dir); var fields=ReadFields(dir);
                                string[] differences=originalFields.TryGetValue(subject,out var expected)
                                    ? expected.Keys.Union(fields.Keys).Where(key=>expected.GetValueOrDefault(key)!=fields.GetValueOrDefault(key)).ToArray()
                                    : ["UnexpectedSubject"];
                                copied.Add(new { Subject=subject, EntryId=(string)((dynamic)item).EntryID, Differences=differences, Matched=differences.Length==0 });
                                allMatched &= differences.Length==0;
                            }
                            finally { Release(item); }
                        }
                    }
                    finally { Release(items); }
                    batches.Add(new { Batch=batch, ProcessSeconds=processTimer.Elapsed.TotalSeconds, VerificationSeconds=verification.Elapsed.TotalSeconds, ExitCode=process.ExitCode, Requested=manifest.Items.Count, TargetCount=Count(target!), Copies=copied });
                    if(process.ExitCode!=0)throw new InvalidOperationException("Cópia nativa retornou falha; lote preservado para revisão, sem repetição.");
                    if(Count(target!)!=manifest.Items.Count)throw new InvalidDataException("Contagem do lote divergiu; sem repetição.");
                    if(!allMatched)break; // Preserve evidence and inspect source before returning review-required.
                }
                finally { Release(target); Release(folders); }
            }
            if(largeBatch && allMatched)
            {
                object? folders=null, seed=null, largeStore=null, largeRoot=null, largeFolders=null, largeTarget=null;
                try
                {
                    folders=((dynamic)root).Folders; seed=((dynamic)folders).Add("Origem artificial de cem itens",6);
                    string seedId=(string)((dynamic)seed).EntryID;
                    var preparation=RunHelper(run,"seed-100",manifest.SourceStoreId,manifest.SourceFolderId,targetStoreId,seedId,manifest.Items.Select(x=>x.EntryId).ToArray(),20);
                    if(Count(seed!)!=100)throw new InvalidDataException("Preparação artificial não gerou exatamente cem itens.");
                    var source100=CaptureHundred(seed!,targetStoreId,seedId,Path.Combine(run,"large-source-before"));
                    var sourceMultiset=source100.Select(x=>x.Hash).OrderBy(x=>x,StringComparer.Ordinal).ToArray();
                    string largePath=Path.Combine(run,"native-batch-100-only.pst");
                    if(File.Exists(largePath))throw new IOException("PST artificial de lote grande já existe.");
                    ((dynamic)session!).AddStoreEx(largePath,2); largeStore=FindStore(session!,largePath) ?? throw new IOException("PST grande não localizado.");
                    largeRoot=((dynamic)largeStore).GetRootFolder(); largeFolders=((dynamic)largeRoot).Folders; largeTarget=((dynamic)largeFolders).Add("Destino artificial de cem itens",6);
                    string largeStoreId=(string)((dynamic)largeStore).StoreID, largeTargetId=(string)((dynamic)largeTarget).EntryID;
                    var transfer=RunHelper(run,"true-batch-100",targetStoreId,seedId,largeStoreId,largeTargetId,source100.Select(x=>x.EntryId).ToArray(),1);
                    var verification=Stopwatch.StartNew();
                    var target100=CaptureHundred(largeTarget!,largeStoreId,largeTargetId,Path.Combine(run,"large-target"));
                    bool multisetMatched=sourceMultiset.SequenceEqual(target100.Select(x=>x.Hash).OrderBy(x=>x,StringComparer.Ordinal));
                    double verificationSeconds=verification.Elapsed.TotalSeconds;
                    var sourceAfter=CaptureHundred(seed!,targetStoreId,seedId,Path.Combine(run,"large-source-after"));
                    bool sourceUnchanged=source100.OrderBy(x=>x.EntryId,StringComparer.Ordinal).SequenceEqual(sourceAfter.OrderBy(x=>x.EntryId,StringComparer.Ordinal));
                    allMatched &= multisetMatched && sourceUnchanged;
                    File.WriteAllText(Path.Combine(run,"true-batch-100-report.json"),JsonSerializer.Serialize(new{SyntheticOnly=true,ProductionEnabled=false,PreparationExcludedFromTransfer=true,Preparation=preparation,Transfer=transfer,VerificationSeconds=verificationSeconds,SourceCount=source100.Count,TargetCount=target100.Count,SourceBytes=source100.Sum(x=>x.Size),MultisetMatched=multisetMatched,SourceUnchanged=sourceUnchanged,DuplicateMultiplicityPreserved=multisetMatched},new JsonSerializerOptions{WriteIndented=true}));
                    Console.WriteLine($"Lote único de cem itens: multiconjunto igual={multisetMatched}, origem preservada={sourceUnchanged}.");
                }
                finally { Release(largeTarget); Release(largeFolders); Release(largeRoot); Release(largeStore); Release(seed); Release(folders); }
            }
            // Paired baseline uses the exact current Export/Import/TargetSnapshot methods,
            // including the durable MSG and stored-copy token Save, but no ledger/removal.
            var baseline=new List<object>();
            using(var backend=new OutlookArchiveBackend())
            {
                var volume=new ArchiveVolume{JobId="native-probe-"+Guid.NewGuid().ToString("N"),Number=1,PstPath=Path.Combine(run,"msg-baseline-only.pst"),SourceStoreId=manifest.SourceStoreId};
                if(File.Exists(volume.PstPath))throw new IOException("Baseline já existe.");
                backend.CreateVolume(volume);
                foreach(var entry in manifest.Items)
                {
                    var item=new ArchiveItem{SourceStoreId=manifest.SourceStoreId,SourceFolderId=manifest.SourceFolderId,SourceEntryId=entry.EntryId,RelativePath=["Baseline artificial"],Token=Guid.NewGuid().ToString("N")};
                    var clock=Stopwatch.StartNew(); var exported=backend.Export(item,Path.Combine(run,$"baseline-export-{entry.Index:D2}")); double exportSeconds=clock.Elapsed.TotalSeconds;
                    item.StageMsgPath=exported.MessagePath; item.StageMsgSha256=exported.MessageSha256;
                    clock.Restart(); var identity=backend.Import(volume,item); double importSeconds=clock.Elapsed.TotalSeconds;
                    clock.Restart(); var targetSnapshot=backend.TargetSnapshot(volume,item,identity,Path.Combine(run,$"baseline-target-{entry.Index:D2}")); double verifySeconds=clock.Elapsed.TotalSeconds;
                    string[] differences=exported.Fields.Keys.Union(targetSnapshot.Fields.Keys).Where(key=>exported.Fields.GetValueOrDefault(key)!=targetSnapshot.Fields.GetValueOrDefault(key)).ToArray();
                    allMatched&=differences.Length==0;
                    baseline.Add(new{entry.Index,ExportSeconds=exportSeconds,ImportSeconds=importSeconds,TargetSnapshotSeconds=verifySeconds,Differences=differences,Matched=differences.Length==0});
                }
            }
            File.WriteAllText(Path.Combine(run,"msg-baseline-report.json"),JsonSerializer.Serialize(new{SyntheticOnly=true,CurrentBackendMethods=true,LedgerIncluded=false,RemovalAllowed=false,Items=baseline},new JsonSerializerOptions{WriteIndented=true}));
            var after=new List<object>(); bool unchanged=Count(source!)==sourceCount;
            foreach(var entry in manifest.Items)
            {
                object? item=null;
                try { item=((dynamic)session!).GetItemFromID(entry.EntryId,manifest.SourceStoreId); GuardItem(item!,manifest.SourceStoreId,manifest.SourceFolderId); bool equal=before[entry.EntryId]==Snapshot(item!,Path.Combine(run,$"source-after-{entry.Index:D2}")); unchanged&=equal; after.Add(new{entry.Index,Unchanged=equal}); }
                finally { Release(item); }
            }
            File.WriteAllText(Path.Combine(run,"source-preservation.json"),JsonSerializer.Serialize(new{SourceCountBefore=sourceCount,SourceCountAfter=Count(source!),Unchanged=unchanged,Items=after,OriginalSnapshotSeconds=beforeSeconds,SourceBytes=sourceBytes}));
            if(!unchanged)throw new InvalidDataException("A origem artificial mudou.");
            if(!allMatched) { error="Diferença nos campos estritos do destino; revisão necessária."; return 3; }
            return 0;
        }
        catch(Exception ex) { error=ex.GetType().Name+": "+ex.Message; Console.Error.WriteLine(error); return 1; }
        finally
        {
            // Do not detach any store: preserve the user's live profile and all evidence.
            Release(root); Release(targetStore); Release(source); Release(sourceStore); Release(session); Release(app);
            if(runCreated && run is not null)File.WriteAllText(Path.Combine(run,"native-probe-report.json"),JsonSerializer.Serialize(new{SyntheticOnly=true,RemovalAllowed=false,ProductionEnabled=false,AllComparedFieldsMatched=allMatched&&error is null,TotalSeconds=total.Elapsed.TotalSeconds,Error=error,Batches=batches},new JsonSerializerOptions{WriteIndented=true}));
        }
    }
    private static string Snapshot(object item,string directory) { Directory.CreateDirectory(directory); string path=Path.Combine(directory,"message.msg"); ((dynamic)item).SaveAs(path,9); return PilotCommand.Snapshot(item,directory,path); }
    private sealed record CapturedItem(string EntryId,string Hash,long Size);
    private static List<CapturedItem> CaptureHundred(object folder,string storeId,string folderId,string directory)
    {
        if(Count(folder)!=100)throw new InvalidDataException("Lote artificial deve conter exatamente cem itens.");
        var result=new List<CapturedItem>(); object? items=null;
        try
        {
            items=((dynamic)folder).Items;
            for(int i=1;i<=100;i++)
            {
                object? item=null;
                try
                {
                    item=((dynamic)items)[i]; GuardItem(item!,storeId,folderId);
                    if(!((string)((dynamic)item).Subject).StartsWith("Teste local Unicode ",StringComparison.Ordinal)||(string)((dynamic)item).To!="fixture@example.invalid")throw new InvalidDataException("Lote contém item não artificial.");
                    string path=Path.Combine(directory,i.ToString("D3")); Snapshot(item!,path);
                    result.Add(new((string)((dynamic)item).EntryID,ArchiveFiles.HashFields(ReadFields(path)),(int)((dynamic)item).Size));
                }
                finally { Release(item); }
            }
        }
        finally { Release(items); }
        if(result.Select(x=>x.EntryId).Distinct(StringComparer.Ordinal).Count()!=100)throw new InvalidDataException("IDs artificiais repetidos.");
        return result;
    }
    private static object RunHelper(string run,string prefix,string sourceStore,string sourceFolder,string targetStore,string targetFolder,string[] ids,int repeats)
    {
        string handoff=Path.Combine(run,prefix+".txt");
        File.WriteAllLines(handoff,new[]{"synthetic-copy-only-v1",sourceStore,sourceFolder,targetStore,targetFolder}.Concat(ids),new UTF8Encoding(false));
        var clock=Stopwatch.StartNew();
        using var process=new Process{StartInfo=new ProcessStartInfo(Path.Combine(AppContext.BaseDirectory,"native-copy.exe")){UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true}};
        process.StartInfo.ArgumentList.Add(handoff); process.StartInfo.ArgumentList.Add(repeats.ToString()); process.Start();
        string stdout=process.StandardOutput.ReadToEnd(),stderr=process.StandardError.ReadToEnd(); process.WaitForExit();
        File.WriteAllText(Path.Combine(run,prefix+"-helper.json"),stdout);File.WriteAllText(Path.Combine(run,prefix+"-helper-errors.txt"),stderr);
        if(process.ExitCode!=0)throw new InvalidOperationException("Helper nativo retornou falha; nenhuma repetição automática.");
        return new{ProcessSeconds=clock.Elapsed.TotalSeconds,Helper=JsonSerializer.Deserialize<JsonElement>(stdout)};
    }
    private static Dictionary<string,string> ReadFields(string directory)=>JsonSerializer.Deserialize<Dictionary<string,string>>(File.ReadAllText(Path.Combine(directory,"field-hashes.json")))!;
    private static int Count(object folder) { object? items=null; try{items=((dynamic)folder).Items;return (int)((dynamic)items).Count;}finally{Release(items);} }
    private static void GuardFolder(object folder,string store,string id) { if((string)((dynamic)folder).StoreID!=store || (string)((dynamic)folder).EntryID!=id)throw new InvalidDataException("Pasta fora do escopo artificial."); }
    private static void GuardItem(object item,string store,string folder) {object? parent=null;try{parent=((dynamic)item).Parent;GuardFolder(parent!,store,folder);if((int)((dynamic)item).Class!=43||(string)((dynamic)item).MessageClass!="IPM.Note")throw new InvalidDataException("Tipo artificial inválido.");}finally{Release(parent);} }
    private static object? FindStore(object session,string path) { object? stores=null;try{stores=((dynamic)session).Stores;for(int i=1;i<=(int)((dynamic)stores).Count;i++){object? store=null;try{store=((dynamic)stores)[i];string candidate=(string)((dynamic)store).FilePath;if(!string.IsNullOrWhiteSpace(candidate)&&Path.GetFullPath(candidate).Equals(path,StringComparison.OrdinalIgnoreCase)){object result=store;store=null;return result;}}finally{Release(store);}}return null;}finally{Release(stores);} }
    private static void Release(object? value) {if(value is not null&&Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
    [DllImport("oleaut32.dll",PreserveSig=true)]private static extern int GetActiveObject(ref Guid id,IntPtr reserved,[MarshalAs(UnmanagedType.IUnknown)]out object? app);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)][return:MarshalAs(UnmanagedType.Bool)]private static extern bool CreateDirectoryW(string path,IntPtr securityAttributes);
    private sealed record FixtureItem(int Index,string Kind,string EntryId);
    private sealed class Manifest { public bool SyntheticOnly{get;set;} public bool SendsEmail{get;set;} public bool RemovalAllowed{get;set;} public string Recipient{get;set;}=""; public string SourcePstPath{get;set;}=""; public string SourceStoreId{get;set;}=""; public string SourceFolderId{get;set;}=""; public List<FixtureItem> Items{get;set;}=[]; }
}
