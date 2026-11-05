using System.Text.Json;

internal sealed partial class OutlookArchiveBackend
{
    internal int DeliveryReportPilot(string root,string jobId,string output)
    {
        string dir=Path.GetFullPath(output);
        if(Directory.Exists(dir))throw new IOException("Use uma pasta nova para o piloto NDR.");Directory.CreateDirectory(dir);
        var created=new List<ArchiveVolume>();
        try
        {
            using var ledger=new ArchiveLedger(Path.Combine(root,"archive.sqlite"),true);
            var reports=new List<ArchiveItem>();
            foreach(var row in ledger.GetItems(jobId).Where(x=>x.State==ArchiveItemState.SkippedUnsupported || x.State==ArchiveItemState.Planned && x.SourceMessageClass==DeliveryReportEvidence.MessageClass).OrderBy(x=>x.EstimatedBytes))
            {
                try{if(row.SourceMessageClass==DeliveryReportEvidence.MessageClass || InspectDeliveryReport(row)?.MessageClass == DeliveryReportEvidence.MessageClass)reports.Add(row);}catch(System.Runtime.InteropServices.COMException){ }
            }
            if(reports.Count==0)throw new ArchiveReviewException("Nenhum aviso NDR do escopo inicial foi localizado.");
            var selected=new[]{reports[0],reports[reports.Count/2],reports[^1]}.DistinctBy(x=>x.Key).ToList();
            var sourceVolume=new ArchiveVolume{JobId=Guid.NewGuid().ToString("N"),Number=1,PstPath=Path.Combine(dir,"ndr-test-source.pst")};
            created.Add(sourceVolume);CreateVolume(sourceVolume);
            var originals=new List<(ArchiveItem Row,ArchiveSnapshot Snapshot)>();DestinationIdentity? fixture=null;
            var fixtures=new List<ArchiveItem>();
            foreach(var row in selected)
            {
                var snapshot=Export(row,Path.Combine(dir,"original-"+row.Key[..12]));originals.Add((row,snapshot));
                fixtures.Add(new ArchiveItem{Key=row.Key,JobId=sourceVolume.JobId,SourceStoreId=row.SourceStoreId,SourceFolderId=row.SourceFolderId,
                    SourceEntryId=row.SourceEntryId,SourceRecordKey=row.SourceRecordKey,SourceMessageClass=DeliveryReportEvidence.MessageClass,
                    StageMsgPath=snapshot.MessagePath,StageMsgSha256=snapshot.MessageSha256,SnapshotHash=snapshot.Hash,SourceChangeKey=snapshot.ChangeKey,
                    RelativePath=["Avisos de falha — teste"]});
            }
            var fixtureBatch=new ArchiveBatch{JobId=sourceVolume.JobId,VolumeNumber=1,SourceStoreId=fixtures[0].SourceStoreId,
                SourceFolderId=fixtures[0].SourceFolderId,RelativePath=fixtures[0].RelativePath,Items=fixtures.Select(x=>x.Key).ToList()};
            fixtureBatch.DirectoryPath=Path.Combine(dir,sourceVolume.JobId,"batches",fixtureBatch.Id);Directory.CreateDirectory(fixtureBatch.DirectoryPath);
            var fixtureFolder=PrepareBatch(sourceVolume,fixtureBatch);fixtureBatch.StoreId=fixtureFolder.StoreId;fixtureBatch.FolderId=fixtureFolder.FolderId;
            CopyBatch(sourceVolume,fixtureBatch,fixtures);
            var candidates=BatchCandidates(sourceVolume,fixtureBatch);var used=new HashSet<string>();
            foreach(var original in originals)
            {
                var item=fixtures.Single(x=>x.Key==original.Row.Key);DestinationIdentity? match=null;
                foreach(var candidate in candidates.Where(x=>!used.Contains(x.EntryId)))
                {
                    var staged=TargetSnapshot(sourceVolume,item,candidate,Path.Combine(dir,"fixture-check-"+Guid.NewGuid().ToString("N")),false);
                    if(RelevantMessageValidation.Equivalent(original.Snapshot,staged)){match=candidate;break;}
                }
                if(match is null)throw new ArchiveReviewException("Cópia nativa de entrada do piloto NDR diverge da origem real.");
                used.Add(match.EntryId);MarkBatchCopy(sourceVolume,fixtureBatch,item,match);
                fixture=FinalizeBatchCopy(sourceVolume,fixtureBatch,item,match);
                object? copied=null,accessor=null;
                try
                {
                    copied=((dynamic)session!).GetItemFromID(fixture.EntryId,fixture.StoreId);GuardBatchCopy(sourceVolume,fixtureBatch,item,copied,false);
                    accessor=((dynamic)copied).PropertyAccessor;((dynamic)accessor).DeleteProperty(TokenProperty);((dynamic)copied).Save();
                }
                finally{Release(accessor);Release(copied);}
                var finalFixture=TargetSnapshot(sourceVolume,item,fixture,Path.Combine(dir,"fixture-final-check-"+item.Key[..12]),false);
                if(!RelevantMessageValidation.Equivalent(original.Snapshot,finalFixture))throw new ArchiveReviewException("A preparação do PST de teste alterou o aviso original.");
            }
            CloseBatch(sourceVolume,fixtureBatch);
            var config=new ArchiveConfig{RootDirectory=Path.Combine(dir,"catalog"),SourceStoreId=fixture!.StoreId,SourceFolderId=fixture.FolderId,
                SyntheticSourcePstPath=sourceVolume.PstPath,SourceLabel="Piloto NDR — cópias isoladas",AllowRetainedMsgBackupForRemoval=true};
            using(var engine=new ArchiveEngine(config.RootDirectory,this))
            {
                string testJob=engine.Plan(config);engine.Copy(testJob);created.AddRange(engine.Volumes(testJob));
                engine.Verify(testJob,allowRetainedBackup:true);
                if(engine.Items(testJob).Count!=selected.Count || engine.Items(testJob).Any(x=>x.State!=ArchiveItemState.Verified))throw new ArchiveReviewException("Piloto NDR: cópias ainda não validadas; consulte o catálogo de teste.");
                engine.Copy(testJob);
                engine.Remove(testJob,false,true);
                if(engine.Items(testJob).Any(x=>x.State!=ArchiveItemState.InDeletedItems))throw new ArchiveReviewException("Piloto NDR: remoção para Itens Excluídos bloqueada.");
                engine.Remove(testJob,true,true);
                if(engine.Items(testJob).Any(x=>x.State!=ArchiveItemState.Removed))throw new ArchiveReviewException("Piloto NDR: exclusão definitiva da cópia de teste bloqueada.");
                foreach(var original in originals)
                {
                    var current=SourceSnapshot(original.Row,null,Path.Combine(dir,"original-recheck-"+original.Row.Key[..12]));
                    if(!RelevantMessageValidation.Equivalent(original.Snapshot,current))throw new ArchiveReviewException("Origem real mudou durante o teste; nenhuma exclusão dela foi solicitada.");
                }
                ArchiveFiles.WriteNew(Path.Combine(dir,"result.json"),new{Passed=true,Reports=selected.Count,Copied=true,Verified=true,MovedToTrash=true,PurgedFixtureOnly=true,RealOriginalsPreserved=true});
                Console.WriteLine($"Piloto NDR aprovado: {selected.Count} avisos, cópia, validação, retomada e exclusão em PSTs isolados. Originais reais preservados.");
            }
            return 0;
        }
        finally
        {
            // Only exact volumes created in this unique pilot directory.
            foreach(var volume in created.DistinctBy(x=>x.PstPath))
            {
                string path=Path.GetFullPath(volume.PstPath);
                if(!path.StartsWith(dir+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new ArchiveReviewException("Limpeza do piloto fora da pasta autorizada bloqueada.");
                object? store=null,folder=null;
                try
                {
                    store=FindStore(path,volume.StoreId);
                    if(store is not null){folder=((dynamic)store).GetRootFolder();((dynamic)session!).RemoveStore(folder);}
                }
                finally{Release(folder);Release(store);}
                attached.Remove(path);tokenIndexes.Remove(path);
                if(pstLeases.Remove(path,out var lease))lease.Dispose();
                try{File.Delete(path);}catch(IOException){Console.WriteLine("PST de teste desanexado, mas o Outlook ainda mantém o arquivo aberto: "+path);}
            }
        }
    }
}
internal static class DeliveryReportPilotAudit
{
    internal static void Run(string directory)
    {
        string root=Path.GetFullPath(directory);
        using var ledger=new ArchiveLedger(Path.Combine(root,"catalog","archive.sqlite"),true);
        var rows=ledger.GetItems(ledger.GetJobs().Single().Id);
        var originals=Directory.GetDirectories(root,"original-*").Where(x=>!Path.GetFileName(x).StartsWith("original-recheck-",StringComparison.Ordinal)).Select(x=>Load(Path.Combine(x,"message.msg"))).ToList();
        var copies=rows.Select(x=>Load(x.StageMsgPath!)).ToList();
        if(rows.Any(x=>x.State!=ArchiveItemState.Removed) || originals.Count!=copies.Count)throw new ArchiveReviewException("Piloto incompleto.");
        bool Match(int i,HashSet<int> used)
        {
            if(i==originals.Count)return true;
            for(int j=0;j<copies.Count;j++)if(!used.Contains(j) && RelevantMessageValidation.Equivalent(originals[i],copies[j]))
            {used.Add(j);if(Match(i+1,used))return true;used.Remove(j);}
            return false;
        }
        bool equal=Match(0,[]);
        ArchiveFiles.WriteNew(Path.Combine(root,"source-import-audit.json"),new{Passed=equal,Messages=originals.Count,ComparedRealSourceToFixture=true,OneToOneMultiplicityPreserved=true,RealOriginalsDeleted=false});
        if(!equal)throw new ArchiveReviewException("A entrada do piloto não equivale integralmente à informação da origem real.");
        Console.WriteLine($"NDR: {originals.Count} cópias de entrada equivalem às origens reais; multiplicidade preservada.");
    }
    private static ArchiveSnapshot Load(string message)
    {
        string dir=Path.GetDirectoryName(message)!;
        return new(ArchiveFiles.ReadFields(Path.Combine(dir,"field-hashes.json")),message,ArchiveFiles.Sha(message),null,RelevantMessageValidation.LoadSidecar(message));
    }
}
