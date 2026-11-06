using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

// Optional, separately versioned evidence. Legacy Fields/Hash remain immutable.
// There is no body-free acceptance path: unsupported profiles retain strict
// comparison and cannot qualify a new native batch.
internal static partial class RelevantMessageValidation
{
    internal const string ProfileVersion = "relevant-message-v1";
    private static readonly object Gate = new();
    private static Process? worker;
    private static bool unavailable;
    private static readonly HashSet<string> RequiredMetadata = new(StringComparer.Ordinal)
    {
        "Meta-SenderName", "Meta-SenderEmailType", "Meta-SenderEmailAddress", "Meta-SentOnBehalfOfName",
        "Meta-SentUtc", "Meta-ReceivedUtc", "Meta-Recipients"
    };
    private static readonly string[] HtmlBodyKeys = ["Text", "Links", "Images", "Resources", "Tables", "Lists", "Preformatted", "ActivePayloads"];
    private static readonly string[] NativeBodyKeys = ["1000001F", "1000001E", "10130102", "1013001F", "10090102"];
    static RelevantMessageValidation() => AppDomain.CurrentDomain.ProcessExit += (_, _) => Stop();

    internal static bool IsAvailable
    {
        get { lock (Gate) { try { return Query(new { Operation = "ping" }).GetProperty("Supported").GetBoolean(); } catch { unavailable = true; Stop(); return false; } } }
    }
    internal static bool Supported(ArchiveSnapshot snapshot) => Supported(snapshot.RelevantFields);
    private static bool HasMetadata(IReadOnlyDictionary<string, string>? fields) => fields is not null && fields.GetValueOrDefault("ProfileVersion") == ProfileVersion &&
        RequiredMetadata.All(key => fields.TryGetValue(key, out string? value) && !string.IsNullOrEmpty(value));
    internal static bool Supported(IReadOnlyDictionary<string, string>? fields)
    {
        if (!HasMetadata(fields) || fields is null) return false;
        IEnumerable<string>? keys = fields.GetValueOrDefault("BodyMode") switch
        {
            "html-relevant-v1" => HtmlBodyKeys,
            "html-controls-v1" => HtmlBodyKeys.Concat(new[] { "Controls" }),
            "html-static-v1" => HtmlBodyKeys.Concat(new[] { "Controls", "InertCSS" }),
            "html-exact-v1" => new[] { "HTML" },
            "signed-mime-exact-v1" => new[] { "SignedMime" },
            "native-exact-v1" or "empty-text-rtf-v1" or "text-derived-rtf-v1" or "text-derived-rtf-v2" or "text-derived-rtf-v3" or "repair-plain-unicode-v1" or "repair-text-representation-v1" or "empty-body-v2" => NativeBodyKeys,
            _ => null
        };
        return keys is not null && keys.All(key => fields.TryGetValue("Body-" + key, out string? value) && !string.IsNullOrEmpty(value));
    }

    internal static bool MetadataEqual(IReadOnlyDictionary<string, string>? left, IReadOnlyDictionary<string, string>? right) =>
        HasMetadata(left) && HasMetadata(right) && RequiredMetadata.All(key => left![key] == right![key]);

    internal static bool OtherInformationExact(IReadOnlyDictionary<string,string> left,IReadOnlyDictionary<string,string> right,
        IReadOnlyDictionary<string,string>? leftProfile,IReadOnlyDictionary<string,string>? rightProfile)
    {
        var pair=SenderAddressEvidence.NormalizePair(new(new Dictionary<string,string>(left),"unused","unused",null,leftProfile is null?null:new Dictionary<string,string>(leftProfile)),new(new Dictionary<string,string>(right),"unused","unused",null,rightProfile is null?null:new Dictionary<string,string>(rightProfile)));
        left=pair.Left.Fields;right=pair.Right.Fields;leftProfile=pair.Left.RelevantFields;rightProfile=pair.Right.RelevantFields;
        if (!Supported(leftProfile) || !Supported(rightProfile) || !MetadataEqual(leftProfile,rightProfile)) return false;
        bool sameSubject=DeliveryReportEvidence.SubjectsEquivalent(left,right,leftProfile,rightProfile);
        var a=left.Where(x=>!(sameSubject && x.Key is "Subject" or "Report-0037001F") && x.Key!="UnRead" && !x.Key.StartsWith("NativeBody-",StringComparison.Ordinal) && !EmbeddedMessageEvidence.IgnoreContainerHash(x.Key,left,right,leftProfile,rightProfile)).ToDictionary(x=>x.Key,x=>x.Value);
        var b=right.Where(x=>!(sameSubject && x.Key is "Subject" or "Report-0037001F") && x.Key!="UnRead" && !x.Key.StartsWith("NativeBody-",StringComparison.Ordinal) && !EmbeddedMessageEvidence.IgnoreContainerHash(x.Key,left,right,leftProfile,rightProfile)).ToDictionary(x=>x.Key,x=>x.Value);
        return Same(a,b);
    }

    internal static bool Equivalent(ArchiveSnapshot left, ArchiveSnapshot right)
    {
        (left,right)=DeliveryReportEvidence.NormalizeSubjects(left,right);
        (left,right)=SenderAddressEvidence.NormalizePair(left,right);
        // Reading a message changes its status, not the archived information.
        // Preserve the immutable evidence; omit only this operational flag in
        // comparisons. Every other field and all body evidence still apply.
        left = left with { Fields = left.Fields.Where(x => x.Key != "UnRead").ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal) };
        right = right with { Fields = right.Fields.Where(x => x.Key != "UnRead").ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal) };
        if (left.Hash == right.Hash)
        {
            bool metadataLeft = HasMetadata(left.RelevantFields), metadataRight = HasMetadata(right.RelevantFields);
            if (!metadataLeft || !metadataRight) return !metadataLeft && !metadataRight;
            return RequiredMetadata.All(key => left.RelevantFields![key] == right.RelevantFields![key]);
        }
        var leftBody=left.Fields.Where(x=>x.Key.StartsWith("NativeBody-",StringComparison.Ordinal)).ToDictionary(x=>x.Key,x=>x.Value);
        var rightBody=right.Fields.Where(x=>x.Key.StartsWith("NativeBody-",StringComparison.Ordinal)).ToDictionary(x=>x.Key,x=>x.Value);
        if (leftBody.Count>0 && Same(leftBody,rightBody) && MetadataEqual(left.RelevantFields,right.RelevantFields))
        {
            var a=left.Fields.Where(x=>!x.Key.StartsWith("NativeBody-",StringComparison.Ordinal) && !EmbeddedMessageEvidence.IgnoreContainerHash(x.Key,left.Fields,right.Fields,left.RelevantFields,right.RelevantFields)).ToDictionary(x=>x.Key,x=>x.Value);
            var b=right.Fields.Where(x=>!x.Key.StartsWith("NativeBody-",StringComparison.Ordinal) && !EmbeddedMessageEvidence.IgnoreContainerHash(x.Key,left.Fields,right.Fields,left.RelevantFields,right.RelevantFields)).ToDictionary(x=>x.Key,x=>x.Value);
            var ap=left.RelevantFields!.Where(x=>!x.Key.StartsWith("Body-",StringComparison.Ordinal) && x.Key is not ("BodyMode" or "ProfileVersion")).ToDictionary(x=>x.Key,x=>x.Value);
            var bp=right.RelevantFields!.Where(x=>!x.Key.StartsWith("Body-",StringComparison.Ordinal) && x.Key is not ("BodyMode" or "ProfileVersion")).ToDictionary(x=>x.Key,x=>x.Value);
            if(Same(a,b) && Same(ap,bp)) return true;
        }
        bool supportedLeft = Supported(left), supportedRight = Supported(right);
        if (!supportedLeft || !supportedRight)
        {
            left = CurrentBodyProfile(left); right = CurrentBodyProfile(right);
            if (!Supported(left) || !Supported(right)) return false;
        }
        // All existing metadata and attachment evidence remains exact. Only the
        // alternate serialized body streams use the independently retained
        // relevant-information profile. Added metadata has to match as well.
        var existingLeft = left.Fields.Where(x => !x.Key.StartsWith("NativeBody-", StringComparison.Ordinal) && !EmbeddedMessageEvidence.IgnoreContainerHash(x.Key, left.Fields, right.Fields, left.RelevantFields, right.RelevantFields)).ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
        var existingRight = right.Fields.Where(x => !x.Key.StartsWith("NativeBody-", StringComparison.Ordinal) && !EmbeddedMessageEvidence.IgnoreContainerHash(x.Key, left.Fields, right.Fields, left.RelevantFields, right.RelevantFields)).ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
        if (!Same(existingLeft, existingRight)) return false;
        if (Same(left.RelevantFields!, right.RelevantFields!)) return true;
        // Retained MSG evidence is immutable. Recompute only its optional body
        // profile using the current parser; never rewrite the original sidecar.
        left = CurrentBodyProfile(left); right = CurrentBodyProfile(right);
        return Supported(left) && Supported(right) && Same(left.RelevantFields!, right.RelevantFields!);
    }
    internal static string DifferenceSummary(ArchiveSnapshot left, ArchiveSnapshot right)
    {
        if (Equivalent(left, right)) return "Nenhuma diferença nas informações relevantes";
        left = CurrentBodyProfile(left); right = CurrentBodyProfile(right);
        var differences = left.Fields.Keys.Union(right.Fields.Keys)
            .Where(key => key != "UnRead" && !key.StartsWith("NativeBody-", StringComparison.Ordinal) &&
                left.Fields.GetValueOrDefault(key) != right.Fields.GetValueOrDefault(key) &&
                !EmbeddedMessageEvidence.IgnoreContainerHash(key, left.Fields, right.Fields, left.RelevantFields, right.RelevantFields)).ToList();
        if (!Supported(left) || !Supported(right)) differences.Add("Formato do corpo ainda não suportado");
        else foreach (string key in left.RelevantFields!.Keys.Union(right.RelevantFields!.Keys))
            if (left.RelevantFields.GetValueOrDefault(key) != right.RelevantFields.GetValueOrDefault(key)) differences.Add(key switch
            {
                "Body-Text" => "Texto do corpo", "Body-Lists" => "Listas do corpo", "Body-Tables" => "Tabelas do corpo",
                "Body-Links" => "Links", "Body-Images" => "Imagens", "Body-Resources" => "Recursos do corpo",
                "Body-HTML" => "HTML integral", "Body-InertCSS" => "CSS e classes de conteúdo gerado", "Body-Controls" => "Campos HTML, valores, rótulos ou posição", "BodyMode" => "Representação do corpo",
                "Meta-Recipients" => "Destinatários, Cc ou Cco", "Meta-SentUtc" => "Data/hora de envio",
                "Meta-ReceivedUtc" => "Data/hora de recebimento", _ when key.StartsWith("Embedded-Item-", StringComparison.Ordinal) => "Conteúdo do e-mail anexado " + key[14..], _ => key
            });
        return differences.Count == 0 ? "Evidência do corpo não equivalente; revisão necessária" : string.Join(", ", differences.Distinct());
    }
    private static ArchiveSnapshot CurrentBodyProfile(ArchiveSnapshot snapshot, bool recoveryBodyContext = false)
    {
        if (!HasMetadata(snapshot.RelevantFields) || !File.Exists(snapshot.MessagePath)) return snapshot;
        try
        {
            if (ArchiveFiles.Sha(snapshot.MessagePath) != snapshot.MessageSha256) return snapshot;
            JsonElement response;
            lock (Gate) response = Query(new { Operation = "profile", Path = Path.GetFullPath(snapshot.MessagePath), ReportEquivalence = recoveryBodyContext || snapshot.RelevantFields?.GetValueOrDefault("Report-BodyProfile") == "ndr-body-v1", RepairEquivalence = recoveryBodyContext });
            if (!response.GetProperty("Supported").GetBoolean()) return snapshot;
            var fields = response.GetProperty("Fields").Deserialize<Dictionary<string, string>>();
            if (fields is null) return snapshot;
            foreach (string key in RequiredMetadata) fields[key] = snapshot.RelevantFields![key];
            if(snapshot.RelevantFields!.TryGetValue("Report-BodyProfile",out string? reportProfile))fields["Report-BodyProfile"]=reportProfile;
            foreach (var entry in snapshot.RelevantFields!.Where(x => x.Key.StartsWith("Embedded-", StringComparison.Ordinal) || x.Key.StartsWith("Report-Subject",StringComparison.Ordinal))) fields[entry.Key] = entry.Value;
            return snapshot with { RelevantFields = fields };
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or InvalidOperationException or JsonException or TimeoutException or System.ComponentModel.Win32Exception) { return snapshot; }
    }
    private static bool Same(IReadOnlyDictionary<string, string> left, IReadOnlyDictionary<string, string> right) =>
        left.Count == right.Count && left.All(x => right.TryGetValue(x.Key, out string? value) && value == x.Value);

    internal static Dictionary<string, string>? Generate(object item, string messagePath, string evidenceDirectory, EmbeddedMessageEvidence.Budget? budget = null, int depth = 0, bool deliveryReportContext = false, bool recoveryBodyContext = false)
    {
        try
        {
            bool report = DeliveryReportEvidence.IsReport(item);
            var effectiveBudget=budget ?? new();
            bool reportContext=report || deliveryReportContext;
            Dictionary<string,string> metadata;
            using(ArchiveOperationTiming.Measure("relevant.read-metadata"))metadata = report ? DeliveryReportEvidence.Metadata(messagePath) : ReadMetadata(item);
            var fields = new Dictionary<string, string> { ["ProfileVersion"] = ProfileVersion, ["BodyMode"] = "unsupported-v1" };
            try
            {
                JsonElement result;
                using(ArchiveOperationTiming.Measure("relevant.python-profile"))lock (Gate) result = Query(new { Operation = "profile", Path = Path.GetFullPath(messagePath), ReportEquivalence = reportContext || recoveryBodyContext, RepairEquivalence = recoveryBodyContext });
                if (result.GetProperty("Supported").GetBoolean())
                    fields = result.GetProperty("Fields").Deserialize<Dictionary<string, string>>() ?? throw new InvalidDataException("Perfil de corpo ausente.");
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or InvalidOperationException or JsonException or System.ComponentModel.Win32Exception or TimeoutException)
            { /* Exact body comparison remains mandatory for this profile. */ }
            if(recoveryBodyContext)foreach(var proof in SenderAddressEvidence.Capture(ArchiveFiles.ReadFields(Path.Combine(evidenceDirectory,"field-hashes.json")),metadata,messagePath))fields[proof.Key]=proof.Value;
            if(reportContext)fields["Report-BodyProfile"]="ndr-body-v1";
            foreach (var field in metadata) fields.Add(field.Key, field.Value);
            try
            {
                foreach (var field in EmbeddedMessageEvidence.Capture(item, evidenceDirectory, effectiveBudget, depth, reportContext, recoveryBodyContext)) fields.Add(field.Key, field.Value);
            }
            catch (Exception ex) when (ex is COMException or InvalidDataException or IOException or InvalidOperationException or System.Reflection.TargetInvocationException)
            {
                ArchiveFiles.WriteNew(Path.Combine(evidenceDirectory, "embedded-status.json"), new { Complete = false, ErrorType = ex.GetType().Name, Reason = ex.Message });
                // Incomplete recursion retains strict container-byte comparison.
            }
            ArchiveFiles.WriteNew(Path.Combine(evidenceDirectory, "relevant-fields.json"), fields);
            return fields;
        }
        catch (Exception ex) when (ex is COMException or InvalidDataException or IOException or InvalidOperationException or JsonException or System.ComponentModel.Win32Exception or TimeoutException)
        {
            // The caller can retain exact validation. A missing profile is never
            // interpreted as equality or used to approve native source removal.
            return null;
        }
    }
    internal static Dictionary<string, string>? LoadSidecar(string messageOrFieldsPath)
    {
        string path = Path.Combine(Path.GetDirectoryName(messageOrFieldsPath)!, "relevant-fields.json");
        if (!File.Exists(path)) return null;
        var fields = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path));
        return HasMetadata(fields) ? fields : null;
    }
    private static Dictionary<string, string> ReadMetadata(object instance)
    {
        dynamic item = instance;
        var raw = new Dictionary<string, object?>
        {
            ["Meta-SenderName"] = (string?)item.SenderName ?? "",
            ["Meta-SenderEmailType"] = (string?)item.SenderEmailType ?? "",
            ["Meta-SenderEmailAddress"] = (string?)item.SenderEmailAddress ?? "",
            ["Meta-SentOnBehalfOfName"] = (string?)item.SentOnBehalfOfName ?? "",
            ["Meta-SentUtc"] = UtcProperty(instance, "0x00390040"),
            ["Meta-ReceivedUtc"] = UtcProperty(instance, "0x0E060040")
        };
        object? recipients = null;
        var rows = new List<object>();
        try
        {
            recipients = item.Recipients;
            int count = (int)((dynamic)recipients).Count;
            for (int index = 1; index <= count; index++)
            {
                object? recipient = null;
                try
                {
                    recipient = ((dynamic)recipients)[index];
                    int type = (int)((dynamic)recipient).Type;
                    if (type is < 1 or > 3) throw new InvalidDataException("Tipo de destinatário não suportado.");
                    rows.Add(new { Type = type, Name = (string?)((dynamic)recipient).Name ?? "", Address = (string?)((dynamic)recipient).Address ?? "",
                        Smtp = OptionalProperty(recipient, "0x39FE001F") as string });
                }
                finally { Release(recipient); }
            }
            if ((int)((dynamic)recipients).Count != count) throw new InvalidDataException("Destinatários mudaram durante a conferência.");
        }
        finally { Release(recipients); }
        raw["Meta-Recipients"] = rows;
        return raw.ToDictionary(x => x.Key, x => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(x.Value)))), StringComparer.Ordinal);
    }
    private static string UtcProperty(object item, string tag)
    {
        object? value = OptionalProperty(item, tag);
        return value is DateTime date ? DateTime.SpecifyKind(date, DateTimeKind.Utc).ToString("O", CultureInfo.InvariantCulture) : "missing";
    }
    private static object? OptionalProperty(object item, string tag)
    {
        object? accessor = null;
        try { accessor = ((dynamic)item).PropertyAccessor; return ((dynamic)accessor).GetProperty("http://schemas.microsoft.com/mapi/proptag/" + tag); }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x8004010F)) { return null; }
        finally { Release(accessor); }
    }
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }

    private static JsonElement Query(object request)
    {
        if (unavailable) throw new InvalidOperationException("Validação relevante indisponível.");
        if (worker is null || worker.HasExited)
        {
            Stop();
            var paths = WorkerPaths();
            var start = new ProcessStartInfo(paths.Python) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8 };
            start.ArgumentList.Add("-I"); start.ArgumentList.Add("-u"); start.ArgumentList.Add(paths.Script); start.ArgumentList.Add(paths.Dependencies);
            worker = Process.Start(start) ?? throw new IOException("Não foi possível iniciar a validação de corpo.");
            worker.ErrorDataReceived += (_, _) => { }; worker.BeginErrorReadLine();
        }
        worker.StandardInput.WriteLine(JsonSerializer.Serialize(request)); worker.StandardInput.Flush();
        var read = worker.StandardOutput.ReadLineAsync();
        if (!read.Wait(TimeSpan.FromSeconds(45))) { Stop(); unavailable = true; throw new TimeoutException("Limite da validação de corpo atingido."); }
        string line = read.Result ?? throw new IOException("O validador de corpo encerrou inesperadamente.");
        if (line.Length > 256 * 1024) throw new InvalidDataException("Resposta de validação inválida.");
        return JsonDocument.Parse(line).RootElement.Clone();
    }
    private static (string Python, string Script, string Dependencies) WorkerPaths()
    {
        string resource = Path.Combine(AppContext.BaseDirectory, "semantic");
        string script = Path.Combine(resource, "semantic_message_worker.py");
        string deps = Path.Combine(resource, "site-packages");
        if (!File.Exists(script))
        {
            DirectoryInfo? directory = new(AppContext.BaseDirectory);
            while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "tools", "semantic_message_worker.py"))) directory = directory.Parent;
            if (directory is null) throw new FileNotFoundException("Recursos da validação de corpo ausentes.");
            script = Path.Combine(directory.FullName, "tools", "semantic_message_worker.py");
            deps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "outlook-relevant-analysis-deps");
        }
        if (!Directory.Exists(deps)) throw new DirectoryNotFoundException("Dependências da validação de corpo ausentes.");
        string configured = Path.Combine(resource, "python.path");
        string python = File.Exists(configured) ? File.ReadAllText(configured).Trim() :
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python314", "python.exe");
        if (!Path.IsPathFullyQualified(python) || !File.Exists(python)) throw new FileNotFoundException("Runtime Python da validação de corpo ausente.");
        return (python, script, deps);
    }
    private static void Stop()
    {
        if (worker is null) return;
        try { if (!worker.HasExited) worker.Kill(true); } catch (InvalidOperationException) { }
        worker.Dispose(); worker = null;
    }
    internal static void SelfTest()
    {
        var fields = new Dictionary<string, string> { ["Subject"] = "same", ["Attachment-1-Hash"] = "same", ["NativeBody-test"] = "one" };
        var relevant = new Dictionary<string, string> { ["ProfileVersion"] = ProfileVersion, ["BodyMode"] = "html-relevant-v1", ["Body-Text"] = "same" };
        foreach (string key in HtmlBodyKeys) relevant["Body-" + key] = "same";
        foreach (string key in RequiredMetadata) relevant[key] = "same";
        ArchiveSnapshot Sample(Dictionary<string,string> values, Dictionary<string,string>? profile) => new(values, "unused.msg", "unused", null, profile);
        var left = Sample(fields, relevant);
        var changedFormat = new Dictionary<string,string>(fields) { ["NativeBody-test"] = "two" };
        if (!Equivalent(left, Sample(changedFormat, new(relevant)))) throw new Exception("Representação equivalente foi rejeitada.");
        foreach (string key in new[] { "Body-Text", "Meta-Recipients", "Meta-SenderEmailAddress", "Meta-SentUtc", "Meta-ReceivedUtc" })
        {
            var changed = new Dictionary<string,string>(relevant) { [key] = "changed" };
            if (Equivalent(left, Sample(changedFormat, changed))) throw new Exception("Informação relevante alterada foi aceita.");
        }
        var exactHtml = new Dictionary<string,string>(relevant);
        foreach (string key in HtmlBodyKeys) exactHtml.Remove("Body-" + key);
        exactHtml["BodyMode"] = "html-exact-v1"; exactHtml["Body-HTML"] = "full-html";
        if (!Equivalent(Sample(fields, exactHtml), Sample(changedFormat, new(exactHtml)))) throw new Exception("HTML integral idêntico foi rejeitado.");
        var differentHtml = new Dictionary<string,string>(exactHtml) { ["Body-HTML"] = "alterado" };
        if (Equivalent(Sample(fields, exactHtml), Sample(changedFormat, differentHtml))) throw new Exception("HTML especial alterado foi aceito.");
        var controls = new Dictionary<string,string>(relevant) { ["BodyMode"] = "html-controls-v1", ["Body-Controls"] = "same-controls" };
        if (!Equivalent(Sample(fields, controls), Sample(changedFormat, new(controls)))) throw new Exception("Campos HTML iguais foram rejeitados.");
        var changedControls = new Dictionary<string,string>(controls) { ["Body-Controls"] = "alterado" };
        if (Equivalent(Sample(fields, controls), Sample(changedFormat, changedControls))) throw new Exception("Campos HTML alterados foram aceitos.");
        changedControls.Remove("Body-Controls");
        if (Supported(changedControls)) throw new Exception("Perfil de campos HTML incompleto foi aceito.");
        var staticHtml = new Dictionary<string,string>(controls) { ["BodyMode"] = "html-static-v1", ["Body-InertCSS"] = "same-css" };
        if (!Equivalent(Sample(fields, staticHtml), Sample(changedFormat, new(staticHtml)))) throw new Exception("CSS inativo preservado foi rejeitado.");
        var changedCss = new Dictionary<string,string>(staticHtml) { ["Body-InertCSS"] = "alterado" };
        if (Equivalent(Sample(fields, staticHtml), Sample(changedFormat, changedCss))) throw new Exception("CSS alterado foi aceito.");
        changedCss.Remove("Body-InertCSS");
        if (Supported(changedCss)) throw new Exception("Perfil CSS incompleto foi aceito.");
        var attachmentChanged = new Dictionary<string,string>(fields) { ["Attachment-1-Hash"] = "changed" };
        if (Equivalent(left, Sample(attachmentChanged, relevant)) || Equivalent(left, Sample(changedFormat, null))) throw new Exception("Evidência incompleta ou anexo alterado foi aceito.");
        if (!Equivalent(Sample(fields, null), Sample(new(fields), null))) throw new Exception("Comparação legada exata alterada.");
        lock (Gate)
        {
            var result = Query(new { Operation = "self-test" });
            if (!result.GetProperty("Supported").GetBoolean() || result.GetProperty("Tests").GetInt32() < 19) throw new Exception("Autoteste do perfil de corpo falhou.");
        }
    }
}
