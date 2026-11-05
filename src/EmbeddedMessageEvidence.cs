using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

// Optional versioned evidence. Original MSG bytes and legacy field hashes are
// never replaced. Only type-5 container hashes may use this complete profile.
internal static class EmbeddedMessageEvidence
{
    private const string Version = "embedded-msg-v1";
    private static readonly string Type5 = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes("5")));
    internal sealed class Budget { internal int Remaining = 64; internal long Bytes = 256L * 1024 * 1024; }
    internal static List<int> Indices(IReadOnlyDictionary<string, string> fields) => fields
        .Where(x => x.Key.StartsWith("Attachment-", StringComparison.Ordinal) && x.Key.EndsWith("-Type", StringComparison.Ordinal) && x.Value == Type5)
        .Select(x => x.Key.Length > 16 && int.TryParse(x.Key[11..^5], NumberStyles.None, CultureInfo.InvariantCulture, out int index) ? index : -1).Order().ToList();
    internal static bool Complete(IReadOnlyDictionary<string, string> fields, IReadOnlyDictionary<string, string>? profile)
    {
        var indices = Indices(fields);
        return indices.Count > 0 && indices.All(index => index > 0 && fields.TryGetValue($"Attachment-{index}-Hash", out var raw) && raw.Length == 64 && raw.All(Uri.IsHexDigit)) && profile is not null && profile.GetValueOrDefault("Embedded-Version") == Version &&
            profile.GetValueOrDefault("Embedded-Indices") == string.Join(",", indices) &&
            profile.Keys.Count(x => x.StartsWith("Embedded-Item-", StringComparison.Ordinal)) == indices.Count &&
            indices.All(index => profile.TryGetValue("Embedded-Item-" + index, out var hash) && hash.Length == 64 && hash.All(Uri.IsHexDigit));
    }
    internal static bool IgnoreContainerHash(string key, IReadOnlyDictionary<string, string> left, IReadOnlyDictionary<string, string> right,
        IReadOnlyDictionary<string, string>? leftProfile, IReadOnlyDictionary<string, string>? rightProfile)
    {
        if (!key.StartsWith("Attachment-", StringComparison.Ordinal) || !key.EndsWith("-Hash", StringComparison.Ordinal)) return false;
        if (!Complete(left, leftProfile) || !Complete(right, rightProfile) || !Indices(left).SequenceEqual(Indices(right))) return false;
        return Indices(left).Any(index => key == $"Attachment-{index}-Hash");
    }
    internal static bool SameExceptContainers(IReadOnlyDictionary<string, string> left, IReadOnlyDictionary<string, string> right,
        IReadOnlyDictionary<string, string>? leftProfile, IReadOnlyDictionary<string, string>? rightProfile, bool ignoreRead = true) =>
        left.Count == right.Count && left.All(x => right.ContainsKey(x.Key) && (ignoreRead && x.Key == "UnRead" ||
            IgnoreContainerHash(x.Key, left, right, leftProfile, rightProfile) || right[x.Key] == x.Value));
    internal static bool SameRepresentations(IReadOnlyDictionary<string, string> left, IReadOnlyDictionary<string, string> right,
        IReadOnlyDictionary<string, string>? leftProfile, IReadOnlyDictionary<string, string>? rightProfile)
    {
        if (!Complete(left, leftProfile) || !Complete(right, rightProfile) || !RelevantMessageValidation.Supported(leftProfile) || !RelevantMessageValidation.Supported(rightProfile)) return false;
        var a = left.Where(x => !x.Key.StartsWith("NativeBody-", StringComparison.Ordinal)).ToDictionary(x => x.Key, x => x.Value);
        var b = right.Where(x => !x.Key.StartsWith("NativeBody-", StringComparison.Ordinal)).ToDictionary(x => x.Key, x => x.Value);
        return SameExceptContainers(a, b, leftProfile, rightProfile);
    }
    internal static bool CanReplay(ArchiveSnapshot original, ArchiveSnapshot replay)
    {
        (original,replay)=SenderAddressEvidence.NormalizePair(original,replay);
        return original.MessageSha256==replay.MessageSha256 && RelevantMessageValidation.MetadataEqual(original.RelevantFields,replay.RelevantFields) && RelevantMessageValidation.Supported(replay) &&
            (Indices(original.Fields).Count==0 && Indices(replay.Fields).Count==0 || Complete(original.Fields,replay.RelevantFields) && Complete(replay.Fields,replay.RelevantFields)) &&
            SameExceptContainers(original.Fields,replay.Fields,replay.RelevantFields,replay.RelevantFields,false);
    }
    internal static string Canonical(ArchiveSnapshot snapshot)
    {
        snapshot=SenderAddressEvidence.Canonical(snapshot);
        if (!RelevantMessageValidation.Supported(snapshot)) throw new InvalidDataException("Corpo do e-mail anexado não suportado.");
        if (Indices(snapshot.Fields).Count > 0 && !Complete(snapshot.Fields, snapshot.RelevantFields))
            throw new InvalidDataException("Evidência recursiva do e-mail anexado incompleta.");
        var data = new Dictionary<string, string> { ["Version"] = Version };
        foreach (var field in snapshot.Fields)
            if (field.Key != "UnRead" && !field.Key.StartsWith("NativeBody-", StringComparison.Ordinal) &&
                !IgnoreContainerHash(field.Key, snapshot.Fields, snapshot.Fields, snapshot.RelevantFields, snapshot.RelevantFields))
                data["Field:" + field.Key] = field.Value;
        foreach (var field in snapshot.RelevantFields!) data["Relevant:" + field.Key] = field.Value;
        return ArchiveFiles.HashFields(data);
    }
    internal static Dictionary<string, string> Capture(object parent, string directory, Budget budget, int depth, bool deliveryReportContext = false, bool recoveryBodyContext = false)
    {
        object? attachments = null, session = null;
        var result = new Dictionary<string, string>(); var indices = new List<int>();
        try
        {
            attachments = ((dynamic)parent).Attachments;
            int count = (int)((dynamic)attachments).Count;
            for (int i = 1; i <= count; i++)
            {
                object? attachment = null, nested = null;
                try
                {
                    attachment = ((dynamic)attachments)[i];
                    if ((int)((dynamic)attachment).Type != 5) continue;
                    if (depth >= 4 || --budget.Remaining < 0) throw new InvalidDataException("Limite da evidência de e-mails anexados.");
                    string serialized = Path.Combine(directory, $"attachment-{i:D3}.msg");
                    if (!File.Exists(serialized)) throw new InvalidDataException("Exportação do e-mail anexado ausente.");
                    budget.Bytes -= new FileInfo(serialized).Length;
                    if (budget.Bytes < 0) throw new InvalidDataException("Limite de bytes de e-mails anexados.");
                    string before = ArchiveFiles.Sha(serialized);
                    string child = Path.Combine(directory, "embedded-" + i.ToString("D3", CultureInfo.InvariantCulture));
                    Directory.CreateDirectory(child);
                    string retained = Path.Combine(child, "message.msg"); File.Copy(serialized, retained, false);
                    string working = Path.Combine(child, "opened.msg"); File.Copy(serialized, working, false);
                    session ??= ((dynamic)parent).Session;
                    nested = ((dynamic)session).OpenSharedItem(working);
                    PilotCommand.Snapshot(nested, child, retained);
                    var fields = ArchiveFiles.ReadFields(Path.Combine(child, "field-hashes.json"));
                    var relevant = RelevantMessageValidation.Generate(nested, retained, child, budget, depth + 1, deliveryReportContext, recoveryBodyContext);
                    if (ArchiveFiles.Sha(serialized) != before || ArchiveFiles.Sha(retained) != before) throw new InvalidDataException("E-mail anexado mudou durante a leitura.");
                    result["Embedded-Item-" + i] = Canonical(new(fields, retained, before, null, relevant)); indices.Add(i);
                }
                finally
                {
                    if (nested is not null) { try { ((dynamic)nested).Close(1); } catch (COMException) { } }
                    Release(nested); Release(attachment);
                }
            }
            if ((int)((dynamic)attachments).Count != count) throw new InvalidDataException("Anexos mudaram durante a leitura.");
            if (indices.Count > 0) { result["Embedded-Version"] = Version; result["Embedded-Indices"] = string.Join(",", indices); }
            return result;
        }
        finally { Release(session); Release(attachments); }
    }
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    internal static void SelfTest()
    {
        string H(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));
        var fields = new Dictionary<string,string> { ["Attachment-1-Type"] = H("5"), ["Attachment-1-Hash"] = H("container-A"),
            ["Attachment-1-Name"] = H("message.msg"), ["Subject"] = H("subject"), ["UnRead"] = H("true"),
            ["MessageClass"] = H("IPM.Note"), ["ReceivedTime"] = H("received"), ["SentOn"] = H("sent"), ["AttachmentCount"] = H("1"),
            ["NativeBody-MSGBodyStream:__substg1.0_1000001F"] = H("body") };
        var profile = new Dictionary<string,string> { ["ProfileVersion"] = "relevant-message-v1", ["BodyMode"] = "html-exact-v1", ["Body-HTML"] = H("body"),
            ["Embedded-Version"] = Version, ["Embedded-Indices"] = "1", ["Embedded-Item-1"] = H("nested content") };
        foreach (string key in new[] { "SenderName", "SenderEmailType", "SenderEmailAddress", "SentOnBehalfOfName", "SentUtc", "ReceivedUtc", "Recipients" }) profile["Meta-" + key] = H(key);
        ArchiveSnapshot S(Dictionary<string,string> f, Dictionary<string,string> p, string? sha = null) => new(f, "unused", sha ?? H("original-msg"), null, p);
        var changedContainer = new Dictionary<string,string>(fields) { ["Attachment-1-Hash"] = H("container-B") };
        void Check(bool value, string error) { if (!value) throw new Exception("E-mail anexado: " + error); }
        Check(RelevantMessageValidation.Equivalent(S(fields,profile), S(changedContainer,new(profile))), "contêiner equivalente rejeitado");
        var changedChild = new Dictionary<string,string>(profile) { ["Embedded-Item-1"] = H("changed body or metadata") };
        Check(!RelevantMessageValidation.Equivalent(S(fields,profile), S(changedContainer,changedChild)), "conteúdo alterado aceito");
        var incomplete = new Dictionary<string,string>(profile); incomplete.Remove("Embedded-Item-1");
        Check(!RelevantMessageValidation.Equivalent(S(fields,profile), S(changedContainer,incomplete)), "recursão incompleta aceita");
        var missingRaw = new Dictionary<string,string>(changedContainer); missingRaw.Remove("Attachment-1-Hash");
        Check(!Complete(missingRaw,profile), "hash bruto ausente aceito");
        var malformed = new Dictionary<string,string> { ["Attachment-Type"] = Type5 };
        Check(!Complete(malformed,profile), "índice de anexo malformado aceito");
        var changedName = new Dictionary<string,string>(changedContainer) { ["Attachment-1-Name"] = H("other.msg") };
        Check(!RelevantMessageValidation.Equivalent(S(fields,profile), S(changedName,new(profile))), "nome alterado aceito");
        var ordinary = new Dictionary<string,string>(fields); ordinary.Remove("Attachment-1-Type");
        var changedOrdinary = new Dictionary<string,string>(ordinary) { ["Attachment-1-Hash"] = H("different ordinary bytes") };
        Check(!RelevantMessageValidation.Equivalent(S(ordinary,profile), S(changedOrdinary,new(profile))), "anexo comum alterado aceito");
        var legacy = profile.Where(x => !x.Key.StartsWith("Embedded-", StringComparison.Ordinal)).ToDictionary(x => x.Key,x=>x.Value);
        Check(CanReplay(S(fields,legacy), S(changedContainer,profile)), "replay íntegro rejeitado");
        Check(!CanReplay(S(fields,legacy), S(changedContainer,profile,H("tampered MSG"))), "MSG alterado aceito");
        var wrongMetadata = new Dictionary<string,string>(profile) { ["Meta-Recipients"] = H("other Cc/Bcc") };
        Check(!CanReplay(S(fields,legacy), S(changedContainer,wrongMetadata)), "destinatários alterados no replay aceitos");
        var wrongBody = new Dictionary<string,string>(changedContainer) { ["NativeBody-MSGBodyStream:__substg1.0_1000001F"] = H("other serialized original") };
        Check(!CanReplay(S(fields,legacy), S(wrongBody,profile)), "campo original alterado no replay aceito");
        Check(Canonical(S(fields,profile)) == Canonical(S(changedContainer,new(profile))), "hash recursivo depende do contêiner");
        foreach (string key in new[] { "Body-HTML", "Meta-Recipients", "Meta-SenderEmailAddress", "Meta-SentUtc", "Meta-ReceivedUtc", "Embedded-Item-1" })
        {
            var changed = new Dictionary<string,string>(profile) { [key] = H("changed") };
            Check(Canonical(S(fields,profile)) != Canonical(S(fields,changed)), "hash recursivo ignorou " + key);
        }
        Check(Canonical(S(fields,profile)) == Canonical(S(new(fields) { ["UnRead"] = H("false") },new(profile))), "leitura alterou conteúdo recursivo");
        ValidationRequest Gate(Dictionary<string,string> sourceProfile, Dictionary<string,string> targetProfile) => new()
        {
            SourceOriginal = fields, SourceCurrent = new Dictionary<string,string>(changedContainer) { ["NativeBody-MSGBodyStream:__substg1.0_1000001F"] = sourceProfile["Body-HTML"] == profile["Body-HTML"] ? fields["NativeBody-MSGBodyStream:__substg1.0_1000001F"] : H("changed parent native") }, TargetBaseline = fields,
            TargetCurrent = new Dictionary<string,string>(fields) { ["Attachment-1-Hash"] = H("container-C") },
            SourceRelevant = profile, SourceCurrentRelevant = sourceProfile, TargetBaselineRelevant = profile, TargetRelevant = targetProfile,
            SourceStoreId = "source-store", SourceEntryId = "source-entry", TargetStoreId = "pst-store", TargetEntryId = "target-entry",
            DestinationIsVerifiedLocalPst = true, DestinationReopenedAndExclusiveRead = true, DestinationBelongsToJob = true,
            SourceStillInExpectedFolder = true, SourceIsSupportedSentMessage = true, StagingMsgHashMatches = true,
            ImportIdentityIsUnique = true, DurablePhaseAllowsRemoval = true
        };
        Check(ArchiveValidation.Evaluate(Gate(new(profile),new(profile))).AllowedToRemove, "validação final rejeitou contêiner equivalente");
        Check(!ArchiveValidation.Evaluate(Gate(changedChild,new(profile))).AllowedToRemove, "validação final aceitou alteração na origem anexada");
        Check(!ArchiveValidation.Evaluate(Gate(new(profile),changedChild)).AllowedToRemove, "validação final aceitou alteração no destino anexado");
        Check(!ArchiveValidation.Evaluate(Gate(incomplete,new(profile))).AllowedToRemove, "validação final aceitou cobertura incompleta");
        var changedParentBody = new Dictionary<string,string>(profile) { ["Body-HTML"] = H("changed parent body") };
        Check(!ArchiveValidation.Evaluate(Gate(changedParentBody,new(profile))).AllowedToRemove, "validação final ignorou mudança no corpo externo");
        Console.WriteLine("E-mails anexados: contêiner, conteúdo, metadados, anexos comuns, cobertura e replay imutável conferidos.");
    }
}
