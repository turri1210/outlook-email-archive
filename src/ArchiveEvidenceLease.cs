using System.Security.Cryptography;

// A cache is valid only while Windows denies writes/deletes to its inputs.
// It lasts for one message, never for the whole operation or a later resume.
internal sealed class ArchiveEvidenceLease : IDisposable
{
    [ThreadStatic] private static ArchiveEvidenceLease? current;
    private readonly ArchiveEvidenceLease? previous;
    private readonly string path;
    private readonly FileStream stream;
    private string? sha;
    private ArchiveSnapshot? backup;
    private bool disposed;
    internal ArchiveEvidenceLease(ArchiveItem item)
    {
        path=Path.GetFullPath(item.StageMsgPath??throw new ArchiveReviewException("Backup original ausente."));
        stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.Read);
        previous=current;current=this;
    }
    internal static bool TryHash(string file,out string hash)
    {
        if(current is {disposed:false} lease && string.Equals(Path.GetFullPath(file),lease.path,StringComparison.OrdinalIgnoreCase))
        {
            if(lease.sha is null){lease.stream.Position=0;lease.sha=Convert.ToHexString(SHA256.HashData(lease.stream));}
            hash=lease.sha;return true;
        }
        hash="";return false;
    }
    internal static bool Guards(string file)=>current is {disposed:false} lease && string.Equals(Path.GetFullPath(file),lease.path,StringComparison.OrdinalIgnoreCase);
    internal static bool TryBackup(ArchiveItem item,out ArchiveSnapshot snapshot)
    {
        if(current is {disposed:false,backup:not null} lease && string.Equals(Path.GetFullPath(item.StageMsgPath!),lease.path,StringComparison.OrdinalIgnoreCase) && lease.backup.MessageSha256==item.StageMsgSha256)
        {snapshot=Clone(lease.backup);return true;}
        snapshot=null!;return false;
    }
    internal static void RememberBackup(ArchiveItem item,ArchiveSnapshot snapshot)
    {
        if(current is {disposed:false} lease && string.Equals(Path.GetFullPath(item.StageMsgPath!),lease.path,StringComparison.OrdinalIgnoreCase) && snapshot.MessageSha256==item.StageMsgSha256)
            lease.backup=Clone(snapshot);
    }
    private static ArchiveSnapshot Clone(ArchiveSnapshot s)=>s with{Fields=new(s.Fields),RelevantFields=s.RelevantFields is null?null:new(s.RelevantFields)};
    public void Dispose(){if(disposed)return;disposed=true;current=previous;stream.Dispose();}
}
