using System.Security.Cryptography;
using System.Text;
using OpenMcdf;

internal sealed partial class OutlookArchiveBackend
{
    internal bool RestoreRepairHtml(ArchiveVolume volume,ArchiveItem row,string evidence)
    {
        if(row.RepairToken is null || row.TargetEntryId is null || row.TargetFolderId is null || row.SourceMessageClass==DeliveryReportEvidence.MessageClass)throw new ArchiveReviewException("Restauração exige cópia de reparo identificada de um e-mail comum.");
        if(ArchiveFiles.Sha(row.StageMsgPath!)!=row.StageMsgSha256)throw new ArchiveReviewException("Backup original alterado.");
        string html;
        using(var root=RootStorage.OpenRead(row.StageMsgPath!))
        {
            if(root.TryOpenStream("__substg1.0_001A001F",out var classStream))
            {
                using(classStream){using var reader=new StreamReader(classStream,Encoding.Unicode,false,1024,true);if(reader.ReadToEnd().TrimEnd('\0')=="IPM.Note.SMIME.MultipartSigned")return false;}
            }
            bool direct=root.TryOpenStream("__substg1.0_10130102",out var stream);
            stream?.Dispose();
            html=direct ? MsgDifferenceAnalysis.ReadHtmlForRestore(row.StageMsgPath!) : RelevantMessageValidation.RecoverEncapsulatedHtml(row.StageMsgPath!);
            // Canonical Unicode composition preserves the text while avoiding
            // Outlook's lossy legacy-codepage fallback for combining accents.
            html=html.Normalize(NormalizationForm.FormC);
        }
        AttachVolume(volume);object? copy=null,parent=null,expected=null,accessor=null;
        try
        {
            copy=((dynamic)session!).GetItemFromID(row.TargetEntryId,volume.StoreId);parent=((dynamic)copy).Parent;expected=TargetFolder(volume,row.RelativePath,false);
            if(expected is null || volume.StoreId==row.SourceStoreId || row.TargetEntryId==row.SourceEntryId || (int)((dynamic)copy).Class!=43 || (string)((dynamic)copy).MessageClass!="IPM.Note" ||
                (string)((dynamic)parent).StoreID!=volume.StoreId || (string)((dynamic)parent).EntryID!=row.TargetFolderId || (string)((dynamic)expected).EntryID!=row.TargetFolderId || (string?)Property(copy,TokenProperty)!=row.DestinationToken)
                throw new ArchiveReviewException("Restauração bloqueada fora da cópia de reparo no PST.");
            Directory.CreateDirectory(evidence);
            var before=Snapshot(copy,Path.Combine(evidence,"before"),true,true);
            ArchiveFiles.WriteNew(Path.Combine(evidence,"intent.json"),new{row.Key,Token=row.DestinationToken,row.TargetEntryId,row.TargetFolderId,volume.StoreId,BackupSha=row.StageMsgSha256,HtmlSha=Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(html))),BeforeMessage=before.MessagePath,BeforeSha=before.MessageSha256,OriginalTouched=false});
            ((dynamic)copy).BodyFormat=2;
            ((dynamic)copy).HTMLBody=html;
            ((dynamic)copy).Save();
            if((string?)Property(copy,TokenProperty)!=row.DestinationToken)throw new ArchiveReviewException("Identificação da cópia mudou durante a restauração.");
            _ = Snapshot(copy,Path.Combine(evidence,"after"),true,true);
            return true;
        }
        finally{Release(accessor);Release(expected);Release(parent);Release(copy);}
    }
}
internal sealed partial class ArchiveEngine
{
    internal void RestoreRepairHtml(string jobId,IReadOnlySet<string> keys)
    {
        var job=ledger.GetJob(jobId);if(backend is not OutlookArchiveBackend outlook)throw new ArchiveReviewException("Restauração HTML exige o Outlook clássico.");
        foreach(var row in ledger.GetItems(jobId).Where(x=>keys.Contains(x.Key)&&x.State==ArchiveItemState.ReviewRequired&&x.RepairToken is not null))
        {
            try
            {
                CheckBackupFile(job,row);_ = OriginalSnapshot(job,row);
                if(outlook.RestoreRepairHtml(Volume(row),row,Evidence(job,row,"restore-html")))Console.WriteLine($"HTML original reaplicado somente à cópia de reparo {row.Key[..12]}; aguardando conferência completa.");
            }
            catch(Exception ex){Console.WriteLine($"Restauração pendente {row.Key[..12]}: {ex.Message}");}
        }
    }
}
