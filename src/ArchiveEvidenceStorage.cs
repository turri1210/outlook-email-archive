using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Microsoft.Win32.SafeHandles;

internal sealed record EvidenceStorageResult(int FilesCompressed,int TemporaryFilesRemoved,int BusyOrUnsupported,long BytesSaved);
internal static class ArchiveEvidenceStorage
{
    [DllImport("kernel32.dll",SetLastError=true)]
    [return:MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeviceIoControl(SafeFileHandle file,uint code,ref ushort input,uint inputSize,IntPtr output,uint outputSize,out uint returned,IntPtr overlapped);
    [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)]
    private static extern uint GetCompressedFileSizeW(string file,out uint high);
    private static long PhysicalSize(string path)
    {
        uint low=GetCompressedFileSizeW(@"\\?\"+Path.GetFullPath(path),out uint high);
        if(low==uint.MaxValue&&Marshal.GetLastWin32Error()!=0)throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        return checked((long)(((ulong)high<<32)|low));
    }
    internal static bool CompressFile(string path,string? expectedSha,out long bytesSaved)
    {
        bytesSaved=0;
        if(!OperatingSystem.IsWindows() || !new DriveInfo(Path.GetPathRoot(Path.GetFullPath(path))!).DriveFormat.Equals("NTFS",StringComparison.OrdinalIgnoreCase))return false;
        var attributes=File.GetAttributes(path);
        if((attributes&(FileAttributes.ReparsePoint|FileAttributes.ReadOnly|FileAttributes.Encrypted))!=0)return false;
        if(attributes.HasFlag(FileAttributes.Compressed))return false;
        using var stream=new FileStream(path,FileMode.Open,FileAccess.ReadWrite,FileShare.Read);
        if(expectedSha is not null)
        {
            if(Convert.ToHexString(SHA256.HashData(stream))!=expectedSha)throw new ArchiveReviewException("Evidência original alterada; compressão bloqueada.");
            stream.Position=0;
        }
        long before=PhysicalSize(path);ushort format=1;
        if(!DeviceIoControl(stream.SafeFileHandle,0x0009C040,ref format,2,IntPtr.Zero,0,out _,IntPtr.Zero))return false;
        stream.Flush(true);
        if(expectedSha is not null)
        {
            stream.Position=0;if(Convert.ToHexString(SHA256.HashData(stream))!=expectedSha)throw new ArchiveReviewException("Integridade da evidência após compressão não confere.");
        }
        bytesSaved=Math.Max(0,before-PhysicalSize(path));return true;
    }
    internal static EvidenceStorageResult Organize(ArchiveJob job,IEnumerable<ArchiveItem> rows,bool move)
    {
        int compressed=0,removed=0,skipped=0;long saved=0;
        foreach(var item in rows.Where(x=>WorkflowPlan.Done(x.State,move) && x.StageMsgPath is not null))
        {
            string directory=ArchiveFiles.Child(job.RootDirectory,Path.Combine(job.RootDirectory,"i",item.Key[..16]));
            if(!Directory.Exists(directory)||File.GetAttributes(directory).HasFlag(FileAttributes.ReparsePoint))continue;
            // Only application-owned temporary copies are disposable. Retained
            // messages, field/relevant sidecars and repair evidence stay intact.
            var options=new EnumerationOptions{RecurseSubdirectories=true,AttributesToSkip=FileAttributes.ReparsePoint,IgnoreInaccessible=false};
            foreach(string file in Directory.EnumerateFiles(directory,"*",options))
            {
                try
                {
                    string resolved=ArchiveFiles.Child(job.RootDirectory,file),name=Path.GetFileName(file);
                    bool original=string.Equals(Path.GetFullPath(file),Path.GetFullPath(item.StageMsgPath!),StringComparison.OrdinalIgnoreCase);
                    bool temporary=name is "opened.msg" or "serialized.msg" || name.Length==21 && name[0]=='w' && name.EndsWith(".msg",StringComparison.OrdinalIgnoreCase) && name.AsSpan(1,16).ToString().All(Uri.IsHexDigit);
                    if(temporary && !original && File.Exists(Path.Combine(Path.GetDirectoryName(file)!,"field-hashes.json")) && !ReferencedByRepair(item,resolved))
                    {
                        // A successfully acquired exclusive handle proves that
                        // Outlook has released this disposable working copy.
                        long bytes=OperatingSystem.IsWindows()?PhysicalSize(file):new FileInfo(file).Length;
                        using(var held=new FileStream(file,FileMode.Open,FileAccess.Read,FileShare.None)){}
                        File.Delete(file);removed++;saved+=bytes;continue;
                    }
                    if(Path.GetExtension(file) is not (".msg" or ".json" or ".txt"))continue;
                    if(CompressFile(file,original?item.StageMsgSha256:null,out long reduction)){compressed++;saved+=reduction;}
                }
                catch(Exception ex) when(ex is IOException or UnauthorizedAccessException or System.ComponentModel.Win32Exception){skipped++;}
            }
        }
        return new(compressed,removed,skipped,saved);
    }
    private static bool ReferencedByRepair(ArchiveItem item,string candidate)
    {
        if(item.RepairManifestPath is null)return false;
        // Unexpected or damaged historical evidence is a reason to retain a
        // temporary copy, not permission to delete it.
        try{return File.ReadAllText(item.RepairManifestPath).Contains(Path.GetFileName(candidate),StringComparison.OrdinalIgnoreCase);}
        catch(IOException){return true;}
    }
}
