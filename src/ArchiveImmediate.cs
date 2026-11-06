internal sealed partial class ArchiveEngine
{
    private void VerifyFreshCopyAndMove(ArchiveJob job,ArchiveItem row,ArchiveVolume volume,ArchiveConfig config,ArchiveSnapshot original,ArchiveSnapshot target)
    {
        using var timing=ArchiveOperationTiming.Measure("inline.verify-and-trash",row.Key[..12]);
        using var lease=new ArchiveEvidenceLease(row);
        // File hashing runs in parallel with the Outlook change-key query.
        // Only immutable local bytes leave the STA; no COM object is shared.
        var hashTask=Task.Run(()=>ArchiveFiles.Sha(row.StageMsgPath!));
        ArchiveSnapshot current;string hash;
        try
        {
            bool unchanged=!string.IsNullOrEmpty(original.ChangeKey) && backend.SourceMatchesChangeKey(row,original.ChangeKey);
            current=unchanged?original:backend.SourceSnapshot(row,null,Evidence(job,row,"inline-source-current"));
        }
        finally{hash=hashTask.GetAwaiter().GetResult();}
        if(hash!=row.StageMsgSha256 || hash!=original.MessageSha256)throw new ArchiveReviewException("Backup MSG da cópia recém-gravada foi alterado.");
        ArchiveEvidenceLease.RememberVerifiedHash(row.StageMsgPath!,hash);
        CheckBackupFile(job,row);
        if(target.Hash!=row.TargetSnapshotHash)throw new ArchiveReviewException("Leitura da cópia recém-gravada difere do checkpoint.");
        var location=backend.LocateSource(row)??throw new ArchiveReviewException("Original não localizado após a cópia; movimentação bloqueada.");
        var proof=EvaluatePrepared(job,row,volume,config,new(false,false),false,location,current,original,target,immediateTrash:true);
        if(!proof.Result.AllowedToRemove){Review(row,string.Join(",",proof.Result.Reasons));return;}
        row.State=ArchiveItemState.Verified;row.Error=null;ledger.UpdateItem(row);
        MoveValidatedOriginal(job,row,proof.Source,proof.Target,location);
        Console.WriteLine($"Fluxo direto: {row.Key[..12]} copiado, conferido e movido para Itens Excluídos na mesma passagem.");
    }
}
