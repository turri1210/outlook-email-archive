using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OpenMcdf;

internal static class SenderAddressEvidence
{
    private const string Version="native-sender-smtp-v1";
    private static readonly string[] Keys=["Sender-AddressProofVersion","Sender-RawFieldHash","Sender-RawMetaHash","Sender-RawTypeHash","Sender-SmtpFieldHash","Sender-SmtpMetaHash"];
    private static string H(string text)=>Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)));
    private static string J(string text)=>H(JsonSerializer.Serialize(text));
    internal static Dictionary<string,string> Capture(IReadOnlyDictionary<string,string> fields,IReadOnlyDictionary<string,string> metadata,string message)
    {
        if(fields.GetValueOrDefault("MessageClass")!=H("IPM.Note"))return new();
        using var root=RootStorage.OpenRead(message);
        string? Read(string tag)
        {
            if(!root.TryOpenStream("__substg1.0_"+tag,out var stream))return null;
            using(stream)
            {
                if(stream.Length>16384)throw new InvalidDataException("Propriedade do remetente excede o limite.");
                byte[] bytes=new byte[checked((int)stream.Length)];stream.ReadExactly(bytes);string value=new UnicodeEncoding(false,false,true).GetString(bytes);
                return value.EndsWith('\0')?value[..^1]:value;
            }
        }
        string? cls=Read("001A001F"),type=Read("0C1E001F"),address=Read("0C1F001F"),smtp=Read("5D01001F");
        if(cls!="IPM.Note" || type is not ("EX" or "SMTP") || address is null || smtp is null || smtp.Length>320 || smtp.Count(c=>c=='@')!=1 || smtp.StartsWith('@') || smtp.EndsWith('@') || smtp.Any(c=>char.IsWhiteSpace(c)||char.IsControl(c)||c is '<' or '>'))return new();
        if(type=="SMTP" && address!=smtp || type=="EX" && !address.StartsWith("/o=",StringComparison.OrdinalIgnoreCase))return new();
        if(fields.GetValueOrDefault("SenderEmailAddress")!=H(address) || metadata.GetValueOrDefault("Meta-SenderEmailAddress")!=J(address) || metadata.GetValueOrDefault("Meta-SenderEmailType")!=J(type))return new();
        return new(){[Keys[0]]=Version,[Keys[1]]=H(address),[Keys[2]]=J(address),[Keys[3]]=J(type),[Keys[4]]=H(smtp),[Keys[5]]=J(smtp)};
    }
    internal static ArchiveSnapshot Bind(ArchiveSnapshot snapshot)
    {
        if(snapshot.RelevantFields is null || snapshot.Fields.GetValueOrDefault("MessageClass")!=H("IPM.Note") || !File.Exists(snapshot.MessagePath))return snapshot;
        if(ArchiveFiles.Sha(snapshot.MessagePath)!=snapshot.MessageSha256)throw new ArchiveReviewException("MSG alterado ao conferir o endereço do remetente.");
        var metadata=new Dictionary<string,string>(snapshot.RelevantFields);
        foreach(var proof in Capture(snapshot.Fields,metadata,snapshot.MessagePath))metadata[proof.Key]=proof.Value;
        return snapshot with{RelevantFields=metadata};
    }
    private static bool Valid(ArchiveSnapshot s)=>s.Fields.GetValueOrDefault("MessageClass")==H("IPM.Note") && s.RelevantFields?.GetValueOrDefault(Keys[0])==Version &&
        s.RelevantFields.GetValueOrDefault(Keys[1])==s.Fields.GetValueOrDefault("SenderEmailAddress") && s.RelevantFields.GetValueOrDefault(Keys[2])==s.RelevantFields.GetValueOrDefault("Meta-SenderEmailAddress") &&
        s.RelevantFields.GetValueOrDefault(Keys[3])==s.RelevantFields.GetValueOrDefault("Meta-SenderEmailType") && Keys.Skip(1).All(key=>s.RelevantFields.TryGetValue(key,out var value)&&value.Length==64&&value.All(Uri.IsHexDigit));
    private static ArchiveSnapshot Normalize(ArchiveSnapshot s)
    {
        var metadata=s.RelevantFields is null?null:new Dictionary<string,string>(s.RelevantFields);
        if(Valid(s))
        {
            s=s with{Fields=new Dictionary<string,string>(s.Fields){["SenderEmailAddress"]=metadata![Keys[4]]}};
            metadata!["Meta-SenderEmailAddress"]=metadata[Keys[5]];metadata["Meta-SenderEmailType"]=J("SMTP");
        }
        if(metadata is not null)foreach(string key in Keys)metadata.Remove(key);
        return s with{RelevantFields=metadata};
    }
    internal static ArchiveSnapshot Canonical(ArchiveSnapshot snapshot)=>Normalize(snapshot);
    internal static (ArchiveSnapshot Left,ArchiveSnapshot Right) NormalizePair(ArchiveSnapshot left,ArchiveSnapshot right)
    {
        if(Valid(left)&&Valid(right)&&left.RelevantFields![Keys[4]]==right.RelevantFields![Keys[4]]&&left.RelevantFields[Keys[5]]==right.RelevantFields[Keys[5]])return(Normalize(left),Normalize(right));
        ArchiveSnapshot Strip(ArchiveSnapshot s)=>s with{RelevantFields=s.RelevantFields?.Where(x=>!Keys.Contains(x.Key,StringComparer.Ordinal)).ToDictionary(x=>x.Key,x=>x.Value)};
        return(Strip(left),Strip(right));
    }
}
