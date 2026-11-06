using System.Security.Cryptography;
using System.Text;

// Pure evidence gate. No Outlook, file access, normalization, mutation or delete.
// The caller must obtain fresh evidence and keep the durable operation lock.
internal static class ArchiveValidation
{
    private static readonly string[] RequiredFields =
    ["MessageClass", "Subject", "ReceivedTime", "SentOn", "AttachmentCount"];

    private static readonly string[] KnownBodyFields =
    [
        "NativeBody-MSGBodyStream:__substg1.0_1000001F",
        "NativeBody-MSGBodyStream:__substg1.0_1000001E",
        "NativeBody-MSGBodyStream:__substg1.0_10130102",
        "NativeBody-MSGBodyStream:__substg1.0_1013001F",
        "NativeBody-MSGBodyStream:__substg1.0_10090102"
    ];

    private static readonly string MissingBodyHash = Hash("missing");
    private static readonly string SupportedMessageClassHash = Hash("IPM.Note");

    internal static ValidationResult Evaluate(ValidationRequest? request)
    {
        if (request is null)
            return new(false, ["validation_request_missing"], false, false, false, "removal-blocked");

        var reasons = new List<string>();
        var original = ValidateSnapshot(request.SourceOriginal, "source_original", reasons);
        var currentSource = ValidateSnapshot(request.SourceCurrent, "source_current", reasons);
        var baseline = ValidateSnapshot(request.TargetBaseline, "target_baseline", reasons);
        var currentTarget = ValidateSnapshot(request.TargetCurrent, "target_current", reasons);

        bool sourceUnchanged = (Exact(original, currentSource) || ContainerExact(original, currentSource, request.SourceRelevant, request.SourceCurrentRelevant)) && Equivalent(original, currentSource, request.SourceRelevant, request.SourceCurrentRelevant);
        bool targetUnchanged = (Exact(baseline, currentTarget) || ContainerExact(baseline, currentTarget, request.TargetBaselineRelevant, request.TargetRelevant)) && (request.TargetBaselineRelevant is null ||
            Equivalent(baseline, currentTarget, request.TargetBaselineRelevant, request.TargetRelevant));
        bool sourceMatchesTarget = Equivalent(original, currentTarget, request.SourceRelevant, request.TargetRelevant);
        Require(sourceUnchanged, "source_changed_or_evidence_invalid");
        Require(targetUnchanged, "target_changed_or_evidence_invalid");
        Require(sourceMatchesTarget, "source_target_fields_differ_or_evidence_invalid");

        bool sourceIdentityKnown = !string.IsNullOrWhiteSpace(request.SourceStoreId) && !string.IsNullOrWhiteSpace(request.SourceEntryId);
        bool targetIdentityKnown = !string.IsNullOrWhiteSpace(request.TargetStoreId) && !string.IsNullOrWhiteSpace(request.TargetEntryId);
        Require(sourceIdentityKnown, "source_identity_missing");
        Require(targetIdentityKnown, "target_identity_missing");
        Require(sourceIdentityKnown && targetIdentityKnown &&
            !string.Equals(request.SourceStoreId, request.TargetStoreId, StringComparison.Ordinal) &&
            !string.Equals(request.SourceEntryId, request.TargetEntryId, StringComparison.Ordinal),
            "source_target_identities_not_distinct");

        Require(request.DestinationIsVerifiedLocalPst, "destination_not_verified_local_pst");
        Require(request.DestinationReopenedAndExclusiveRead || (request.IndependentMsgBackupValidated &&
            (request.DestinationReopenedAndRevalidated || request.DestinationStoredItemReadBack && request.TrashOnly)), "destination_durable_reopen_or_independent_backup_missing");
        Require(request.DestinationBelongsToJob, "destination_not_owned_by_job");
        Require(request.SourceStillInExpectedFolder, "source_not_in_expected_folder");
        Require(request.SourceIsSupportedSentMessage, "source_not_supported_sent_message");
        Require(request.StagingMsgHashMatches, "staging_message_hash_mismatch_or_missing");
        Require(request.ImportIdentityIsUnique, "destination_identity_not_unique");
        Require(request.DurablePhaseAllowsRemoval, "durable_phase_does_not_allow_removal");

        bool allowed = reasons.Count == 0;
        return new(allowed, reasons.ToArray(), sourceUnchanged, targetUnchanged, sourceMatchesTarget,
            allowed ? "information-evidence-and-identity-validated" : "removal-blocked");

        void Require(bool condition, string reason) { if (!condition) reasons.Add(reason); }
    }

    // Copy into an ordinal dictionary: a case-insensitive input comparer must not
    // make an incorrectly cased required field appear to have been captured.
    private static Dictionary<string, string>? ValidateSnapshot(
        IReadOnlyDictionary<string, string>? snapshot, string label, List<string> reasons)
    {
        if (snapshot is null || snapshot.Count == 0)
        { reasons.Add(label + "_snapshot_missing_or_empty"); return null; }

        var fields = new Dictionary<string, string>(StringComparer.Ordinal);
        bool valid = true;
        foreach (var field in snapshot)
        {
            if (string.IsNullOrWhiteSpace(field.Key) || !IsSha256(field.Value) || !fields.TryAdd(field.Key, field.Value))
                valid = false;
        }
        if (!valid) reasons.Add(label + "_snapshot_has_invalid_hash_or_field");
        foreach (string required in RequiredFields)
        {
            if (!fields.ContainsKey(required))
            { valid = false; reasons.Add(label + "_required_field_missing:" + required); }
        }
        if (fields.TryGetValue("MessageClass", out var messageClass) &&
            !string.Equals(messageClass, SupportedMessageClassHash, StringComparison.Ordinal) && messageClass != Hash(DeliveryReportEvidence.MessageClass) && messageClass != Hash(DeliveryReportEvidence.AutomaticReplyClass))
        { valid = false; reasons.Add(label + "_message_class_not_supported"); }

        if (fields.GetValueOrDefault("MessageClass") == Hash(DeliveryReportEvidence.MessageClass) &&
            (fields.GetValueOrDefault("Report-Version") != Hash("ndr-v1") || DeliveryReportEvidence.RequiredFields.Any(key=>!fields.ContainsKey(key))))
        { valid=false; reasons.Add(label+"_delivery_report_evidence_missing"); }

        // Missing streams are represented by SHA256("missing") in field hashes.
        // Identical absences or an arbitrary NativeBody-* name are no body proof.
        bool recognizedBodyPresent = KnownBodyFields.Any(key =>
            fields.TryGetValue(key, out string? value) && IsSha256(value) &&
            !string.Equals(value, MissingBodyHash, StringComparison.OrdinalIgnoreCase));
        if (!recognizedBodyPresent)
        { valid = false; reasons.Add(label + "_recognized_present_body_evidence_missing"); }
        return valid ? fields : null;
    }

    private static bool ContainerExact(Dictionary<string, string>? left, Dictionary<string, string>? right,
        IReadOnlyDictionary<string, string>? leftRelevant, IReadOnlyDictionary<string, string>? rightRelevant) =>
        left is not null && right is not null && RelevantMessageValidation.OtherInformationExact(left,right,leftRelevant,rightRelevant);
    private static bool Exact(Dictionary<string, string>? left, Dictionary<string, string>? right) =>
        left is { Count: > 0 } && right is not null && left.Count(x => x.Key != "UnRead") == right.Count(x => x.Key != "UnRead") &&
        left.Where(x => x.Key != "UnRead").All(x => right.TryGetValue(x.Key, out string? value) && string.Equals(x.Value, value, StringComparison.Ordinal));
    private static bool Equivalent(Dictionary<string, string>? left, Dictionary<string, string>? right,
        IReadOnlyDictionary<string, string>? leftRelevant, IReadOnlyDictionary<string, string>? rightRelevant) =>
        left is not null && right is not null && RelevantMessageValidation.Equivalent(
            new(left, "unused", "unused", null, leftRelevant?.ToDictionary(x => x.Key, x => x.Value)),
            new(right, "unused", "unused", null, rightRelevant?.ToDictionary(x => x.Key, x => x.Value)));

    private static bool IsSha256(string? value) => value is { Length: 64 } && value.All(char.IsAsciiHexDigit);
    private static string Hash(string value) => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value)));

    internal static void SelfTest()
    {
        ValidationRequest Good()
        {
            var fields = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["MessageClass"] = Hash("IPM.Note"), ["Subject"] = Hash("Assunto artificial — ação"),
                ["ReceivedTime"] = Hash("2026-10-26T12:00:00.0000000"),
                ["SentOn"] = Hash("2026-10-26T11:59:00.0000000"), ["AttachmentCount"] = Hash("1"),
                ["Attachment-1-Hash"] = Hash(Hash("bytes artificiais")),
                [KnownBodyFields[0]] = Hash("length=10;sha256=" + Hash("corpo")),
                [KnownBodyFields[4]] = Hash("length=20;sha256=" + Hash("RTF"))
            };
            return new ValidationRequest
            {
                SourceOriginal = fields, SourceCurrent = fields, TargetBaseline = fields, TargetCurrent = fields,
                SourceStoreId = "source-store", SourceEntryId = "source-entry",
                TargetStoreId = "target-store", TargetEntryId = "target-entry",
                DestinationIsVerifiedLocalPst = true, DestinationReopenedAndExclusiveRead = true,
                DestinationBelongsToJob = true, SourceStillInExpectedFolder = true, SourceIsSupportedSentMessage = true,
                StagingMsgHashMatches = true, ImportIdentityIsUnique = true, DurablePhaseAllowsRemoval = true
            };
        }

        static Dictionary<string, string> Copy(IReadOnlyDictionary<string, string> input) =>
            input.ToDictionary(x => x.Key, x => x.Value, StringComparer.Ordinal);
        static void Block(ValidationRequest input, string test)
        { if (Evaluate(input).AllowedToRemove) throw new Exception("Remoção indevidamente permitida: " + test); }

        var good = Good();
        if (!Evaluate(good).AllowedToRemove) throw new Exception("Evidência completa válida foi bloqueada.");
        var immediate=good with {DestinationReopenedAndExclusiveRead=false,DestinationReopenedAndRevalidated=false,IndependentMsgBackupValidated=true,DestinationStoredItemReadBack=true,TrashOnly=true};
        if(!Evaluate(immediate).AllowedToRemove)throw new Exception("Leitura recém-gravada com backup MSG íntegro foi bloqueada para Itens Excluídos.");
        Block(immediate with{TrashOnly=false},"leitura direta para exclusão definitiva");
        Block(immediate with{IndependentMsgBackupValidated=false},"leitura direta sem backup MSG íntegro");
        Block(immediate with{StagingMsgHashMatches=false},"leitura direta com MSG alterado");
        var automatic=Copy(good.SourceOriginal!);automatic["MessageClass"]=Hash(DeliveryReportEvidence.AutomaticReplyClass);
        var automaticRequest=good with {SourceOriginal=automatic,SourceCurrent=automatic,TargetBaseline=automatic,TargetCurrent=automatic};
        if(!Evaluate(automaticRequest).AllowedToRemove)throw new Exception("Resposta automática íntegra bloqueada.");
        var wrongAutomatic=Copy(automatic);wrongAutomatic["Subject"]=Hash("alterado");
        Block(automaticRequest with {TargetCurrent=wrongAutomatic},"resposta automática alterada");
        wrongAutomatic=Copy(automatic);wrongAutomatic["MessageClass"]=Hash("IPM.Note.Rules.Unknown");
        Block(automaticRequest with {TargetCurrent=wrongAutomatic},"subclasse não suportada");
        var ndr=Copy(good.SourceOriginal!);ndr["MessageClass"]=Hash(DeliveryReportEvidence.MessageClass);
        foreach(string key in DeliveryReportEvidence.RequiredFields)ndr[key]=Hash(key=="Report-Version"?"ndr-v1":key);
        var report=good with {SourceOriginal=ndr,SourceCurrent=ndr,TargetBaseline=ndr,TargetCurrent=ndr};
        if(!Evaluate(report).AllowedToRemove)throw new Exception("NDR íntegro bloqueado.");
        foreach(string key in DeliveryReportEvidence.RequiredFields)
        {
            var changed=Copy(ndr);changed.Remove(key);Block(report with {TargetCurrent=changed},"NDR sem prova: "+key);
            changed=Copy(ndr);changed[key]=Hash("alterado");Block(report with {SourceCurrent=changed},"NDR alterado: "+key);
        }
        Dictionary<string,string> ReportProfile(Dictionary<string,string> fields,string normalized)
        {
            var p=new Dictionary<string,string>{["ProfileVersion"]="relevant-message-v1",["BodyMode"]="native-exact-v1"};
            foreach(string key in new[]{"SenderName","SenderEmailType","SenderEmailAddress","SentOnBehalfOfName","SentUtc","ReceivedUtc","Recipients"})p["Meta-"+key]=Hash(key);
            foreach(string key in new[]{"1000001F","1000001E","10130102","1013001F","10090102"})p["Body-"+key]=Hash(key);
            p["Report-SubjectProofVersion"]="ndr-native-subject-v1";p["Report-SubjectRawHash"]=fields["Subject"];p["Report-SubjectNormalizedHash"]=Hash(normalized);return p;
        }
        var portuguese=Copy(ndr);portuguese["Subject"]=portuguese["Report-0037001F"]=Hash("unicode:Não é possível entregar: Assunto original");
        var english=Copy(ndr);english["Subject"]=english["Report-0037001F"]=Hash("unicode:Undeliverable: Assunto original");
        var pt=ReportProfile(portuguese,"unicode:Assunto original");var en=ReportProfile(english,"unicode:Assunto original");
        var translated=report with {SourceOriginal=portuguese,SourceCurrent=english,TargetBaseline=portuguese,TargetCurrent=portuguese,SourceRelevant=pt,SourceCurrentRelevant=en,TargetRelevant=pt,TargetBaselineRelevant=pt};
        if(!Evaluate(translated).AllowedToRemove)throw new Exception("Prefixo nativo NDR traduzido bloqueou conteúdo equivalente.");
        var changedSubject=new Dictionary<string,string>(en){["Report-SubjectNormalizedHash"]=Hash("outro assunto")};
        Block(translated with {SourceCurrentRelevant=changedSubject},"Assunto NDR realmente alterado");
        changedSubject=new(en);changedSubject.Remove("Report-SubjectRawHash");Block(translated with {SourceCurrentRelevant=changedSubject},"Tradução NDR sem vinculação ao assunto nativo");
        var noReportProof=Copy(good.SourceOriginal!);noReportProof["MessageClass"]=Hash(DeliveryReportEvidence.MessageClass);
        Block(good with {SourceOriginal=noReportProof,SourceCurrent=noReportProof,TargetBaseline=noReportProof,TargetCurrent=noReportProof},"classe NDR sem conferência do aviso");
        foreach (string field in new[] { "Subject", "ReceivedTime", "Attachment-1-Hash", KnownBodyFields[0] })
        {
            var changed = Copy(good.SourceCurrent!); changed[field] = Hash("alterado");
            Block(good with { SourceCurrent = changed }, "origem mudou: " + field);
            changed = Copy(good.TargetCurrent!); changed[field] = Hash("alterado");
            Block(good with { TargetCurrent = changed }, "destino mudou: " + field);
        }
        var newField = Copy(good.TargetCurrent!); newField["AdditionalField"] = Hash("novo");
        Block(good with { TargetCurrent = newField }, "campo adicional no destino");
        Block(good with { TargetBaseline = null }, "baseline de destino ausente");
        Block(good with { SourceOriginal = new Dictionary<string, string>() }, "evidência vazia");
        Block(good with { TargetStoreId = good.SourceStoreId }, "mesmo armazenamento");
        Block(good with { TargetEntryId = good.SourceEntryId }, "mesmo EntryID");
        Block(good with { SourceEntryId = " " }, "identidade ausente");
        Block(good with { ImportIdentityIsUnique = false }, "cópias duplicadas");
        Block(good with { StagingMsgHashMatches = false }, "MSG intermediário alterado");
        Block(good with { DestinationIsVerifiedLocalPst = false }, "destino não verificado");
        Block(good with { DestinationReopenedAndExclusiveRead = false }, "sem evidência de disco");
        if (!Evaluate(good with { DestinationReopenedAndExclusiveRead = false, IndependentMsgBackupValidated = true, DestinationReopenedAndRevalidated = true }).AllowedToRemove)
            throw new Exception("Backup MSG independente validado e destino reaberto foram bloqueados.");
        Block(good with { DestinationReopenedAndExclusiveRead = false, IndependentMsgBackupValidated = true }, "backup sem reabertura");
        Block(good with { DestinationReopenedAndExclusiveRead = false, DestinationReopenedAndRevalidated = true }, "reabertura sem backup independente");
        Block(good with { DestinationBelongsToJob = false }, "destino de outro trabalho");
        Block(good with { SourceStillInExpectedFolder = false }, "origem fora da pasta");
        Block(good with { SourceIsSupportedSentMessage = false }, "rascunho/não suportado");
        Block(good with { DurablePhaseAllowsRemoval = false }, "estado persistente não permite");

        var missing = Copy(good.SourceOriginal!);
        foreach (string key in KnownBodyFields) missing[key] = MissingBodyHash;
        Block(good with { SourceOriginal = missing, SourceCurrent = missing, TargetBaseline = missing, TargetCurrent = missing }, "corpos todos ausentes");

        var fakeBody = Copy(good.SourceOriginal!);
        foreach (string key in KnownBodyFields) fakeBody.Remove(key);
        fakeBody["NativeBody-unrecognized"] = Hash("fake");
        Block(good with { SourceOriginal = fakeBody, SourceCurrent = fakeBody, TargetBaseline = fakeBody, TargetCurrent = fakeBody }, "nome de corpo não reconhecido");

        var wrongCase = Copy(good.SourceOriginal!);
        wrongCase.Remove("MessageClass"); wrongCase["messageclass"] = SupportedMessageClassHash;
        Block(good with { SourceOriginal = wrongCase, SourceCurrent = wrongCase, TargetBaseline = wrongCase, TargetCurrent = wrongCase }, "nome de campo com caixa incorreta");
        var badHash = Copy(good.SourceOriginal!); badHash["Subject"] = "not-a-sha256";
        Block(good with { SourceOriginal = badHash, SourceCurrent = badHash, TargetBaseline = badHash, TargetCurrent = badHash }, "hash inválido");
        if (Evaluate(null).AllowedToRemove) throw new Exception("Pedido ausente permitiu remoção.");
        var relevant = new Dictionary<string, string> { ["ProfileVersion"] = RelevantMessageValidation.ProfileVersion, ["BodyMode"] = "html-relevant-v1" };
        foreach (string key in new[] { "Text", "Links", "Images", "Resources", "Tables", "Lists", "Preformatted", "ActivePayloads" }) relevant["Body-" + key] = Hash(key);
        foreach (string key in new[] { "SenderName", "SenderEmailType", "SenderEmailAddress", "SentOnBehalfOfName", "SentUtc", "ReceivedUtc", "Recipients" }) relevant["Meta-" + key] = Hash(key);
        var reformatted = Copy(good.TargetCurrent!); reformatted[KnownBodyFields[0]] = Hash("outra representação preservada");
        var semantic = good with { TargetBaseline = reformatted, TargetCurrent = reformatted, SourceRelevant = relevant,
            SourceCurrentRelevant = relevant, TargetBaselineRelevant = relevant, TargetRelevant = relevant };
        if (!Evaluate(semantic).AllowedToRemove) throw new Exception("Informações relevantes iguais com outra representação foram bloqueadas.");
        foreach (string key in new[] { "Body-Text", "Body-Links", "Meta-Recipients", "Meta-SentUtc", "Meta-ReceivedUtc", "Meta-SenderEmailAddress" })
        {
            var changed = new Dictionary<string, string>(relevant) { [key] = Hash("informação alterada") };
            Block(semantic with { TargetBaselineRelevant = changed, TargetRelevant = changed }, "informação relevante alterada: " + key);
        }
        Block(semantic with { TargetRelevant = null }, "perfil do corpo ausente");
        Console.WriteLine("Validação: identidades, disco e informações relevantes conferidos; corpo/data/participantes/anexos alterados e ambiguidades bloqueiam remoção.");
    }
}

internal sealed record ValidationRequest
{
    public IReadOnlyDictionary<string, string>? SourceRelevant { get; init; }
    public IReadOnlyDictionary<string, string>? SourceCurrentRelevant { get; init; }
    public IReadOnlyDictionary<string, string>? TargetRelevant { get; init; }
    public IReadOnlyDictionary<string, string>? TargetBaselineRelevant { get; init; }
    public IReadOnlyDictionary<string, string>? SourceOriginal { get; init; }
    public IReadOnlyDictionary<string, string>? SourceCurrent { get; init; }
    public IReadOnlyDictionary<string, string>? TargetCurrent { get; init; }
    public IReadOnlyDictionary<string, string>? TargetBaseline { get; init; }
    public string? SourceStoreId { get; init; }
    public string? SourceEntryId { get; init; }
    public string? TargetStoreId { get; init; }
    public string? TargetEntryId { get; init; }
    public bool DestinationIsVerifiedLocalPst { get; init; }
    public bool DestinationReopenedAndExclusiveRead { get; init; }
    public bool IndependentMsgBackupValidated { get; init; }
    public bool DestinationReopenedAndRevalidated { get; init; }
    public bool DestinationStoredItemReadBack {get;init;}
    public bool TrashOnly {get;init;}
    public bool DestinationBelongsToJob { get; init; }
    public bool SourceStillInExpectedFolder { get; init; }
    public bool SourceIsSupportedSentMessage { get; init; }
    public bool StagingMsgHashMatches { get; init; }
    public bool ImportIdentityIsUnique { get; init; }
    public bool DurablePhaseAllowsRemoval { get; init; }
}

internal sealed record ValidationResult(
    bool AllowedToRemove,
    IReadOnlyList<string> Reasons,
    bool SourceUnchanged,
    bool TargetUnchanged,
    bool SourceMatchesTarget,
    string Classification);
