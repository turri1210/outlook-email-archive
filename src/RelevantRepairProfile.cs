using System.Security.Cryptography;
using System.Text;

internal static partial class RelevantMessageValidation
{
    internal static string RecoverEncapsulatedHtml(string path)
    {
        System.Text.Json.JsonElement result;
        lock(Gate) result=Query(new { Operation="recover-encapsulated-html", Path=Path.GetFullPath(path) });
        if(!result.GetProperty("Supported").GetBoolean())throw new ArchiveReviewException("HTML encapsulado não pôde ser recuperado integralmente.");
        return result.GetProperty("Html").GetString() ?? throw new ArchiveReviewException("HTML recuperado ausente.");
    }
    // An explicit repair can reinterpret body presentation from immutable MSG
    // bytes. The old sidecars remain intact; metadata/attachments stay bound.
    internal static ArchiveSnapshot RecoveryProfile(ArchiveSnapshot snapshot) => RecoveryProfile(snapshot,new EmbeddedMessageEvidence.Budget(),0);
    private static ArchiveSnapshot RecoveryProfile(ArchiveSnapshot snapshot,EmbeddedMessageEvidence.Budget budget,int depth)
    {
        if(!HasMetadata(snapshot.RelevantFields) || !File.Exists(snapshot.MessagePath))return snapshot;
        if(ArchiveFiles.Sha(snapshot.MessagePath)!=snapshot.MessageSha256)throw new ArchiveReviewException("MSG alterado ao conferir equivalência do reparo.");
        // Fake providers serialize JSON; they must retain their existing proof.
        using(var stream=File.OpenRead(snapshot.MessagePath))
        { Span<byte> magic=stackalloc byte[8];if(stream.Read(magic)!=8 || !magic.SequenceEqual(new byte[]{0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1}))return snapshot; }
        foreach(var field in MsgBodyFingerprint.Read(snapshot.MessagePath))
        {
            string hash=Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(field.Value)));
            if(snapshot.Fields.GetValueOrDefault("NativeBody-"+field.Key)!=hash)throw new ArchiveReviewException("Corpo retido não corresponde ao snapshot original do reparo.");
        }
        var updated=SenderAddressEvidence.Bind(CurrentBodyProfile(snapshot,true));
        if(!Supported(updated))return snapshot;
        var indices=EmbeddedMessageEvidence.Indices(snapshot.Fields);
        if(indices.Count==0 || !EmbeddedMessageEvidence.Complete(snapshot.Fields,snapshot.RelevantFields))return updated;
        var profile=new Dictionary<string,string>(updated.RelevantFields!);
        foreach(int index in indices)
        {
            if(depth>=4 || --budget.Remaining<0)throw new ArchiveReviewException("Limite da comparação recursiva do reparo.");
            string folder=Path.Combine(Path.GetDirectoryName(snapshot.MessagePath)!,"embedded-"+index.ToString("D3",System.Globalization.CultureInfo.InvariantCulture));
            string message=Path.Combine(folder,"message.msg");
            if(!File.Exists(message))throw new ArchiveReviewException("Evidência retida do e-mail anexado ausente.");
            budget.Bytes-=new FileInfo(message).Length;if(budget.Bytes<0)throw new ArchiveReviewException("Limite de bytes da comparação recursiva do reparo.");
            string sha=ArchiveFiles.Sha(message);
            if(snapshot.Fields.GetValueOrDefault($"Attachment-{index}-Hash")!=Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sha))))throw new ArchiveReviewException("E-mail anexado retido não corresponde ao contêiner registrado.");
            var fields=ArchiveFiles.ReadFields(Path.Combine(folder,"field-hashes.json"));
            var relevant=LoadSidecar(message) ?? throw new ArchiveReviewException("Metadados retidos do e-mail anexado ausentes.");
            var child=new ArchiveSnapshot(fields,message,sha,null,relevant);
            if(EmbeddedMessageEvidence.Canonical(child)!=snapshot.RelevantFields!.GetValueOrDefault("Embedded-Item-"+index))throw new ArchiveReviewException("Evidência do e-mail anexado difere do registro original.");
            profile["Embedded-Item-"+index]=EmbeddedMessageEvidence.Canonical(RecoveryProfile(child,budget,depth+1));
        }
        return updated with{RelevantFields=profile};
    }
}
