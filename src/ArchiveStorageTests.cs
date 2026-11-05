using System.Security.Cryptography;
using Microsoft.Data.Sqlite;

internal static class ArchiveStorageTests
{
    private static void Check(bool value,string reason){if(!value)throw new Exception(reason);}
    internal static int Run()
    {
        string root=Path.Combine(Path.GetTempPath(),"archive-storage-"+Guid.NewGuid().ToString("N")[..12]);Directory.CreateDirectory(root);
        Index();Lease(root);History(root);Compression(root);
        Console.WriteLine("Armazenamento: índice único, cache com bloqueio de escrita, histórico íntegro e compressão transparente aprovados.");return 0;
    }
    private static void Index()
    {
        var index=new SourceLookupIndex();for(int i=0;i<10000;i++)index.Add(new("entry-"+i,i.ToString("X16"),"token-"+i));
        for(int i=0;i<10000;i++)Check(index.ByRecord(i.ToString("x16"))?.EntryId=="entry-"+i,"Índice retornou outra mensagem.");
        index.Add(new("duplicate","0000000000000001","other"));
        try{index.ByRecord("0000000000000001");throw new Exception("Ambiguidade aceita.");}catch(ArchiveReviewException){}
        index.Remove("entry-1");Check(index.ByRecord("0000000000000001")?.EntryId=="duplicate","Remoção do índice deixou registro antigo.");
        index.Add(new("duplicate","AABB","replacement"));Check(index.ByToken("other") is null&&index.ByRecord("aabb")?.Token=="replacement","Índice reutilizou metadados antigos.");
    }
    private static void Lease(string root)
    {
        string file=Path.Combine(root,"immutable.msg");File.WriteAllText(file,"original data");string sha=ArchiveFiles.Sha(file);var row=new ArchiveItem{StageMsgPath=file,StageMsgSha256=sha};
        using(var lease=new ArchiveEvidenceLease(row))
        {
            Check(ArchiveFiles.Sha(file)==sha&&ArchiveFiles.Sha(file)==sha,"Cache mudou a assinatura.");
            try{using var write=new FileStream(file,FileMode.Open,FileAccess.Write,FileShare.ReadWrite);throw new Exception("Cache aceitou escrita concorrente.");}catch(IOException){}
            ArchiveEvidenceLease.RememberBackup(row,new(new(){{"field","original"}},file,sha));
            Check(ArchiveEvidenceLease.TryBackup(row,out var copy),"Replay não foi reutilizado.");copy.Fields["field"]="changed";
            Check(ArchiveEvidenceLease.TryBackup(row,out var again)&&again.Fields["field"]=="original","Cache expôs um snapshot mutável.");
        }
        Check(!ArchiveEvidenceLease.TryBackup(row,out _),"Cache sobreviveu à mensagem.");File.WriteAllText(file,"changed data");Check(ArchiveFiles.Sha(file)!=sha,"Assinatura antiga sobreviveu ao bloqueio.");
    }
    private static void History(string root)
    {
        string path=Path.Combine(root,"catalog.sqlite");string job=Guid.NewGuid().ToString("N");
        using(var ledger=new ArchiveLedger(path))
        {
            ledger.InsertJob(new(){Id=job,RootDirectory=root});
            for(int i=0;i<240;i++)
            {
                var row=new ArchiveItem{JobId=job,SourceStoreId="synthetic-source",SourceEntryId="source-"+i};ledger.AddItem(row);
                row.Error=new string('X',24000)+" — evento "+i;ledger.UpdateItem(row);
            }
            var original=ledger.GetEvents(job);ledger.ReclaimHistorySpace();long before=new FileInfo(path).Length;
            Check(ledger.PackHistory(job,200)>0&&ledger.GetEvents(job).SequenceEqual(original),"Primeira compactação alterou o histórico.");
            Check(ledger.PackHistory(job,20)>0&&ledger.GetEvents(job).SequenceEqual(original),"Retomada da compactação perdeu eventos.");
            ledger.ReclaimHistorySpace();Check(new FileInfo(path).Length<before,"Compactação não recuperou espaço.");
        }
        using(var read=new ArchiveLedger(path,true))Check(read.GetEvents(job).Count==481,"Leitura posterior perdeu eventos arquivados.");
        using(var sql=new SqliteConnection("Data Source="+path)){sql.Open();using var command=sql.CreateCommand();command.CommandText="UPDATE event_archives SET data=zeroblob(length(data)) WHERE first_id=(SELECT MIN(first_id) FROM event_archives)";command.ExecuteNonQuery();}
        using(var damaged=new ArchiveLedger(path,true))try{damaged.GetEvents(job);throw new Exception("Histórico corrompido aceito.");}catch(Exception ex) when(ex is IOException or InvalidDataException){}
    }
    private static void Compression(string root)
    {
        string jobRoot=Path.Combine(root,"job");Directory.CreateDirectory(jobRoot);string key=new string('A',64);
        string directory=Path.Combine(jobRoot,"i",key[..16],"original");Directory.CreateDirectory(directory);
        string file=Path.Combine(directory,"message.msg");File.WriteAllText(file,string.Concat(Enumerable.Repeat("Mensagem artificial preservada.\r\n",60000)));string sha=ArchiveFiles.Sha(file);
        File.WriteAllText(Path.Combine(directory,"field-hashes.json"),"{}");File.WriteAllText(Path.Combine(directory,"opened.msg"),"Disposable working copy");
        string protectedPst=Path.Combine(jobRoot,"protected.pst");File.WriteAllText(protectedPst,"Do not compress this PST");var attributes=File.GetAttributes(protectedPst);
        var row=new ArchiveItem{Key=key,State=ArchiveItemState.InDeletedItems,StageMsgPath=file,StageMsgSha256=sha};
        var result=ArchiveEvidenceStorage.Organize(new(){RootDirectory=jobRoot},[row],true);
        Check(ArchiveFiles.Sha(file)==sha,"Compressão mudou os bytes da mensagem.");Check(!File.Exists(Path.Combine(directory,"opened.msg")),"Cópia temporária não foi removida.");
        Check(File.GetAttributes(protectedPst)==attributes,"Compressão alcançou o PST.");
        row.State=ArchiveItemState.ReviewRequired;File.WriteAllText(Path.Combine(directory,"opened.msg"),"Evidence of pending item");
        ArchiveEvidenceStorage.Organize(new(){RootDirectory=jobRoot},[row],true);Check(File.Exists(Path.Combine(directory,"opened.msg")),"Pendência perdeu evidência temporária.");
        Console.WriteLine($"Evidência artificial: {result.FilesCompressed} arquivos comprimidos; {result.BytesSaved} bytes físicos recuperados.");
    }
}
