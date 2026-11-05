using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OpenMcdf;

// ReportItem has no MailItem sender/date/recipient getters. Read its native
// report metadata and recipient table from the retained MSG instead.
internal static class DeliveryReportEvidence
{
    internal const string MessageClass = "REPORT.IPM.Note.NDR";
    internal const string AutomaticReplyClass = "IPM.Note.Rules.OofTemplate.Microsoft";
    internal static bool SupportedMailClass(string messageClass) => messageClass is "IPM.Note" or AutomaticReplyClass;
    internal static bool IgnoredMeeting(string? messageClass) => messageClass is "IPM.Schedule.Meeting.Request" or "IPM.Schedule.Meeting.Resp.Pos" or "IPM.Schedule.Meeting.Canceled";
    internal static bool IsReport(object item) => (int)((dynamic)item).Class == 46 && (string)((dynamic)item).MessageClass == MessageClass;
    internal static bool SupportedSource(object item) => IsReport(item) ||
        (int)((dynamic)item).Class == 43 && SupportedMailClass((string)((dynamic)item).MessageClass) && (bool)((dynamic)item).Sent;
    private static readonly string[] RootTags = ["001A001F","0037001F","00390040","0E060040","0C1A001F","0C1E001F","0C1F001F","5D01001F",
        "0042001F","0064001F","0065001F","5D02001F","0E02001F","0E03001F","0E04001F","007D001F","00310003","00320040",
        "0049001F","00510102","0054001F","00550040","0C040003","0C050003","0C06000B","1001001F","1035001F","1039001F"];
    // Identity/row handles may be reassigned in the destination store. They
    // convey no report text/address/status; recipient order remains explicit.
    private static readonly HashSet<ushort> RecipientHandles = [0x0FF6,0x0FF9,0x0FFA,0x0FFB,0x0FFF,0x3000];
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
    internal static Dictionary<string,string> Read(string path)
    {
        if(new FileInfo(path).Length>256L*1024*1024)throw new InvalidDataException("Limite de arquivo NDR excedido.");
        using var root=RootStorage.OpenRead(path);
        var props=ReadProperties(root,32);
        if (ReadValue(root,props,0x001A001F) != "unicode:"+MessageClass) throw new InvalidDataException("Classe NDR não confirmada no MSG.");
        var values=new Dictionary<string,string>(StringComparer.Ordinal) { ["Report-Version"]="ndr-v1" };
        foreach(string tag in RootTags) values["Report-"+tag]=ReadValue(root,props,uint.Parse(tag,NumberStyles.HexNumber,CultureInfo.InvariantCulture));
        var recipients=root.EnumerateEntries().Where(x=>x.Name.StartsWith("__recip_version1.0_#",StringComparison.Ordinal)).OrderBy(x=>x.Name,StringComparer.Ordinal).ToList();
        if(recipients.Count>10000) throw new InvalidDataException("Limite de destinatários do aviso excedido.");
        byte[] header=StreamBytes(root,"__properties_version1.0")!;
        if(BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(16,4))!=recipients.Count) throw new InvalidDataException("Tabela de destinatários do aviso incompleta.");
        var recipientHashes=new Dictionary<string,string>();
        for(int i=0;i<recipients.Count;i++)
        {
            var child=root.OpenStorage(recipients[i].Name);
            var entries=ReadProperties(child,8);var data=new Dictionary<string,string>();
            foreach(uint tag in entries.Keys.Order())
            {
                ushort id=(ushort)(tag>>16);
                if(RecipientHandles.Contains(id))continue;
                if(id>=0x8000)throw new InvalidDataException("Propriedade nomeada de destinatário NDR ainda não suportada.");
                data[tag.ToString("X8")]=Hash(ReadValue(child,entries,tag));
            }
            // Include streams as well: an orphaned/unlisted diagnostic must
            // not disappear unnoticed. Only identifier streams are excluded.
            foreach(var entry in child.EnumerateEntries())
            {
                if(entry.Name=="__properties_version1.0")continue;
                if(!entry.Name.StartsWith("__substg1.0_",StringComparison.Ordinal) || entry.Name.Length!=20 || !uint.TryParse(entry.Name[12..],NumberStyles.HexNumber,CultureInfo.InvariantCulture,out uint tag))
                    throw new InvalidDataException("Estrutura de destinatário NDR não suportada.");
                if(!RecipientHandles.Contains((ushort)(tag>>16)) && !entries.ContainsKey(tag))
                    throw new InvalidDataException("Stream de destinatário sem propriedade registrada.");
            }
            recipientHashes[i.ToString("D5")]=ArchiveFiles.HashFields(data);
        }
        values["Report-RecipientCount"]=recipients.Count.ToString(CultureInfo.InvariantCulture);
        values["Report-Recipients"]=ArchiveFiles.HashFields(recipientHashes);
        return values;
    }
    internal static IEnumerable<string> RequiredFields => RootTags.Select(x=>"Report-"+x).Concat(new[]{"Report-Version","Report-RecipientCount","Report-Recipients"});
    internal static Dictionary<string,string> Metadata(string path)
    {
        var data=Read(path);
        var result = new Dictionary<string,string>
        {
            ["Meta-SenderName"]=Hash(data["Report-0C1A001F"]),["Meta-SenderEmailType"]=Hash(data["Report-0C1E001F"]),
            ["Meta-SenderEmailAddress"]=Hash(data["Report-0C1F001F"]),["Meta-SentOnBehalfOfName"]=Hash(data["Report-0042001F"]),
            ["Meta-SentUtc"]=Hash(data["Report-00390040"]),["Meta-ReceivedUtc"]=Hash(data["Report-0E060040"]),["Meta-Recipients"]=Hash(data["Report-Recipients"])
        };
        foreach(var proof in SubjectProof(path))result[proof.Key]=proof.Value;
        return result;
    }
    private static readonly string[] SubjectProofKeys=["Report-SubjectProofVersion","Report-SubjectRawHash","Report-SubjectNormalizedHash"];
    private static Dictionary<string,string> SubjectProof(string path)
    {
        using var root=RootStorage.OpenRead(path);var props=ReadProperties(root,32);
        string raw=ReadValue(root,props,0x0037001F), normalized=ReadValue(root,props,0x0E1D001F), prefix=ReadValue(root,props,0x003D001F);
        if(ReadValue(root,props,0x001A001F)!="unicode:"+MessageClass || !raw.StartsWith("unicode:",StringComparison.Ordinal) || !normalized.StartsWith("unicode:",StringComparison.Ordinal) || !prefix.StartsWith("unicode:",StringComparison.Ordinal))return new();
        string label=prefix[8..];
        if(label is not ("Não é possível entregar: " or "Undeliverable: ") || raw!="unicode:"+label+normalized[8..])return new();
        return new(){[SubjectProofKeys[0]]="ndr-native-subject-v1",[SubjectProofKeys[1]]=Hash(raw),[SubjectProofKeys[2]]=Hash(normalized)};
    }
    internal static ArchiveSnapshot BindSubjectProof(ArchiveSnapshot snapshot)
    {
        if(snapshot.Fields.GetValueOrDefault("MessageClass")!=Hash(MessageClass) || snapshot.RelevantFields is null || !File.Exists(snapshot.MessagePath))return snapshot;
        if(ArchiveFiles.Sha(snapshot.MessagePath)!=snapshot.MessageSha256)throw new ArchiveReviewException("MSG alterado ao conferir o assunto do aviso.");
        var proof=SubjectProof(snapshot.MessagePath);if(proof.Count==0)return snapshot;
        var relevant=new Dictionary<string,string>(snapshot.RelevantFields);foreach(var pair in proof)relevant[pair.Key]=pair.Value;
        return snapshot with{RelevantFields=relevant};
    }
    private static bool ValidSubjectProof(IReadOnlyDictionary<string,string> fields,IReadOnlyDictionary<string,string>? profile) =>
        fields.GetValueOrDefault("MessageClass")==Hash(MessageClass) && fields.GetValueOrDefault("Report-Version")==Hash("ndr-v1") &&
        profile?.GetValueOrDefault(SubjectProofKeys[0])=="ndr-native-subject-v1" && profile.GetValueOrDefault(SubjectProofKeys[1])==fields.GetValueOrDefault("Subject") &&
        profile.GetValueOrDefault(SubjectProofKeys[1])==fields.GetValueOrDefault("Report-0037001F") && profile.TryGetValue(SubjectProofKeys[2],out string? normalized) && normalized.Length==64 && normalized.All(Uri.IsHexDigit);
    internal static bool SubjectsEquivalent(IReadOnlyDictionary<string,string> left,IReadOnlyDictionary<string,string> right,IReadOnlyDictionary<string,string>? a,IReadOnlyDictionary<string,string>? b) =>
        ValidSubjectProof(left,a) && ValidSubjectProof(right,b) && a![SubjectProofKeys[2]]==b![SubjectProofKeys[2]];
    internal static (ArchiveSnapshot Left,ArchiveSnapshot Right) NormalizeSubjects(ArchiveSnapshot left,ArchiveSnapshot right)
    {
        if(SubjectsEquivalent(left.Fields,right.Fields,left.RelevantFields,right.RelevantFields))
        {
            string hash=left.RelevantFields![SubjectProofKeys[2]];
            var a=new Dictionary<string,string>(left.Fields){["Subject"]=hash,["Report-0037001F"]=hash};
            var b=new Dictionary<string,string>(right.Fields){["Subject"]=hash,["Report-0037001F"]=hash};
            left=left with{Fields=a};right=right with{Fields=b};
        }
        Dictionary<string,string>? Clean(Dictionary<string,string>? values)=>values?.Where(x=>!SubjectProofKeys.Contains(x.Key,StringComparer.Ordinal)).ToDictionary(x=>x.Key,x=>x.Value);
        return(left with{RelevantFields=Clean(left.RelevantFields)},right with{RelevantFields=Clean(right.RelevantFields)});
    }
    private static Dictionary<uint,byte[]> ReadProperties(Storage storage,int header)
    {
        byte[] bytes=StreamBytes(storage,"__properties_version1.0") ?? throw new InvalidDataException("Propriedades NDR ausentes.");
        if(bytes.Length<header || (bytes.Length-header)%16!=0 || bytes.Length>4*1024*1024)throw new InvalidDataException("Tabela de propriedades NDR inválida.");
        var result=new Dictionary<uint,byte[]>();
        for(int p=header;p<bytes.Length;p+=16)
            if(!result.TryAdd(BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(p,4)),bytes.AsSpan(p+8,8).ToArray()))throw new InvalidDataException("Propriedade NDR duplicada.");
        return result;
    }
    private static string ReadValue(Storage storage,Dictionary<uint,byte[]> entries,uint tag)
    {
        if(!entries.TryGetValue(tag,out byte[]? value))return "missing";
        ushort type=(ushort)tag;
        if(type is 0x001F or 0x001E or 0x0102)
        {
            byte[] bytes=StreamBytes(storage,"__substg1.0_"+tag.ToString("X8")) ?? throw new InvalidDataException("Valor NDR ausente.");
            if(type==0x001F)
            {
                string text=new UnicodeEncoding(false,false,true).GetString(bytes);
                return "unicode:"+(text.EndsWith('\0')?text[..^1]:text);
            }
            return "bytes:"+Convert.ToHexString(SHA256.HashData(bytes))+":"+bytes.Length;
        }
        int length=type switch {0x0002 or 0x000B=>2,0x0003 or 0x0004=>4,0x0005 or 0x0006 or 0x0007 or 0x0014 or 0x0040=>8,_=>throw new InvalidDataException("Tipo de propriedade NDR ainda não suportado: "+tag.ToString("X8"))};
        return "fixed:"+Convert.ToHexString(value.AsSpan(0,length));
    }
    private static byte[]? StreamBytes(Storage storage,string name)
    {
        if(!storage.TryOpenStream(name,out var stream))return null;
        using(stream)
        {
            if(stream.Length>32*1024*1024)throw new InvalidDataException("Limite de stream NDR excedido.");
            byte[] bytes=new byte[checked((int)stream.Length)];stream.ReadExactly(bytes);return bytes;
        }
    }
}

internal interface IArchiveDeliveryReportBackend
{
    SourceDescriptor? InspectDeliveryReport(ArchiveItem item);
}
internal sealed partial class OutlookArchiveBackend : IArchiveDeliveryReportBackend
{
    public SourceDescriptor? InspectDeliveryReport(ArchiveItem item)
    {
        object? source=null;
        try
        {
            source=((dynamic)session!).GetItemFromID(item.SourceEntryId,item.SourceStoreId);
            if(!DeliveryReportEvidence.IsReport(source))
            {
                string messageClass=(string)((dynamic)source).MessageClass;
                if(!DeliveryReportEvidence.IgnoredMeeting(messageClass) && messageClass!=DeliveryReportEvidence.AutomaticReplyClass)return null;
                object? parent=null;
                try
                {
                    parent=((dynamic)source).Parent;
                    if(!SameIdentity((string)((dynamic)parent).StoreID,item.SourceStoreId) || !SameIdentity((string)((dynamic)parent).EntryID,item.SourceFolderId) || item.SourceRecordKey is not null && BinaryProperty(source,"0x0FF90102")!=item.SourceRecordKey)
                        throw new ArchiveReviewException("Item de reunião fora da pasta inicial; classificação preservada.");
                    bool automatic=messageClass==DeliveryReportEvidence.AutomaticReplyClass && (int)((dynamic)source).Class==43;
                    DateTime? ReadMailDate(string tag)=>Property(source,"http://schemas.microsoft.com/mapi/proptag/"+tag) is DateTime d ? DateTime.SpecifyKind(d,DateTimeKind.Utc):null;
                    var(mailDate,mailBasis)=automatic ? ArchiveDateFilter.SelectDate(ReadMailDate("0x0E060040"),ReadMailDate("0x00390040")) : (null,(string?)null);
                    return new(item.SourceStoreId,item.SourceFolderId,item.SourceEntryId,item.RelativePath,item.EstimatedBytes,automatic,automatic && (bool)((dynamic)source).Sent,item.SourceRecordKey,mailDate,mailBasis,messageClass);
                }
                finally{Release(parent);}
            }
            Release(source);source=null;
            source=OpenSource(item,item.SourceEntryId,item.SourceFolderId,item.SourceRecordKey);
            DateTime? ReadDate(string tag)=>Property(source,"http://schemas.microsoft.com/mapi/proptag/"+tag) is DateTime d ? DateTime.SpecifyKind(d,DateTimeKind.Utc):null;
            var(date,basis)=ArchiveDateFilter.SelectDate(ReadDate("0x0E060040"),ReadDate("0x00390040"));
            return new(item.SourceStoreId,item.SourceFolderId,item.SourceEntryId,item.RelativePath,(int)((dynamic)source).Size,true,true,item.SourceRecordKey,date,basis,DeliveryReportEvidence.MessageClass);
        }
        finally{Release(source);}
    }
}
internal sealed partial class ArchiveEngine
{
    internal void PrepareDeliveryReports(string jobId)
    { var job=ledger.GetJob(jobId); var config=Envelope(job).Config; SealPlan(job); UpgradeDeliveryReports(job,config); Status(jobId); }
    private void UpgradeDeliveryReports(ArchiveJob job,ArchiveConfig config)
    {
        if(backend is not IArchiveDeliveryReportBackend reports)return;
        int count=0, ignored=0;
        foreach(var row in WorkItems(job.Id).Where(x=>x.State==ArchiveItemState.SkippedUnsupported))
        {
            try
            {
                var source=reports.InspectDeliveryReport(row);
                if(source is not null && DeliveryReportEvidence.IgnoredMeeting(source.MessageClass))
                {
                    row.SourceMessageClass=source.MessageClass;row.State=ArchiveItemState.SkippedByPolicy;row.Error=null;ledger.UpdateItem(row);ignored++;continue;
                }
                if(source is null || !source.IsSupported || !source.IsSent || !ArchiveDateFilter.Includes(source.MessageDateUtc,config.CutoffDateInclusive,config.CutoffTimeZoneId))continue;
                row.SourceMessageClass=source.MessageClass;row.State=ArchiveItemState.Planned;row.Error=null;
                ledger.UpdateItem(row);count++;
            }
            catch(Exception ex) when(ex is System.Runtime.InteropServices.COMException or ArchiveReviewException)
            { Console.WriteLine($"Aviso ainda não habilitado: {row.Key[..12]} — {ex.Message}"); }
        }
        if(ignored>0)Console.WriteLine($"Convites, aceites e cancelamentos desconsiderados por sua escolha: {ignored}. Originais preservados.");
        if(count>0)Console.WriteLine($"Avisos de falha e respostas automáticas habilitados no escopo inicial: {count}. Nenhuma mensagem nova foi incluída.");
    }
}
