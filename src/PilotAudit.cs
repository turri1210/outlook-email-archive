using System.Text.Json;

// Offline reconciliation only: never connects to Outlook, retries a copy or
// assumes that an event in a log proves the message still exists in a store.
internal static class PilotAudit
{
    internal static int Execute(string directory)
    {
        try
        {
            var result = Read(directory);
            Console.WriteLine(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
            return result.ReviewRequired ? 3 : 0;
        }
        catch (Exception ex) { Console.Error.WriteLine("Auditoria indisponível: " + ex.Message); return 1; }
    }

    internal static AuditResult Read(string directory)
    {
        string run = Path.GetFullPath(directory);
        var result = new AuditResult { RunDirectory = run };
        var events = ReadJournal(Path.Combine(run, "journal.jsonl"), result.Errors);
        PilotReport? report = null;
        string reportPath = Path.Combine(run, "report.json");
        if (File.Exists(reportPath))
        {
            try { report = JsonSerializer.Deserialize<PilotReport>(File.ReadAllText(reportPath)); }
            catch (JsonException) { result.Errors.Add("Relatório final inválido/truncado."); }
        }
        if (report is null)
        {
            result.Errors.Add("Relatório final indisponível; somente diário pode ser auditado.");
            var planned = events.FirstOrDefault(x => State(x) == "planned" && x.TryGetProperty("Report", out _));
            if (planned.ValueKind == JsonValueKind.Object)
                report = planned.GetProperty("Report").Deserialize<PilotReport>();
            report ??= new PilotReport { Status = "report-unavailable" };
        }
        result.RecordedStatus = report.Status;
        result.RecordedSourceCountBefore = report.SourceCountBefore;
        result.RecordedSourceCountAfter = report.SourceCountAfter;
        result.JournalImportedDestinationIds = events.Where(x => State(x) == "imported" && x.TryGetProperty("DestinationId", out _))
            .Select(x => x.GetProperty("DestinationId").GetString() ?? "").ToArray();
        var copyStarted = Indices(events, "source-copy-started");
        var copyCreated = Indices(events, "source-copy-created");
        var importStarted = Indices(events, "import-started");
        var imported = Indices(events, "imported");
        result.CopyCreationAmbiguous = copyStarted.Except(copyCreated).Order().ToArray();
        result.ImportAmbiguous = importStarted.Except(imported).Order().ToArray();
        result.TemporaryCopiesPossiblyRemaining = copyCreated.Except(imported).Order().ToArray();
        result.MetadataRestoreAmbiguous = Indices(events, "unsent-delivery-time-restore-started")
            .Except(Indices(events, "unsent-delivery-time-restored")).Order().ToArray();
        int index = 0;
        foreach (var item in report.Items)
        {
            index++;
            string itemRoot = Path.Combine(run, $"item-{index:D3}");
            result.Items.Add(new AuditItem
            {
                Index = index, RecordedStatus = item.Status, SourceId = item.SourceId,
                RecordedDestinationId = item.DestinationId, TemporarySourceCopyId = item.TemporarySourceCopyId,
                DestinationComparison = Compare(Path.Combine(itemRoot, "source", "field-hashes.json"),
                    Path.Combine(itemRoot, "destination", "field-hashes.json")),
                SourceAfterComparison = Compare(Path.Combine(itemRoot, "source", "field-hashes.json"),
                    Path.Combine(itemRoot, "source-after", "field-hashes.json"))
            });
        }
        foreach (string recovery in Directory.EnumerateDirectories(run, "reconciliation-*", SearchOption.TopDirectoryOnly).Order(StringComparer.Ordinal))
        {
            if (File.GetAttributes(recovery).HasFlag(FileAttributes.ReparsePoint))
            { result.Errors.Add("Reconciliação com redirecionamento ignorada: " + Path.GetFileName(recovery)); continue; }
            var recoveryEvents = ReadJournal(Path.Combine(recovery, "journal.jsonl"), result.Errors);
            int begun = recoveryEvents.Count(x => State(x) == "destination-relocation-started");
            int finished = recoveryEvents.Count(x => State(x) == "destination-relocated");
            if (begun != finished) result.Errors.Add("Relocação ambígua em " + Path.GetFileName(recovery));
            foreach (var entry in recoveryEvents.Where(x => State(x) == "destination-relocated"))
                result.DestinationRelocations.Add(new(Path.GetFileName(recovery),
                    entry.GetProperty("PreviousId").GetString()!, entry.GetProperty("DestinationId").GetString()!));
            // A completed move log only establishes recorded identity transitions.
            // It never promotes an earlier failed content comparison to verified.
        }
        result.ReviewRequired = report.Status != "pilot-completed" || result.Errors.Count != 0 ||
            result.CopyCreationAmbiguous.Length != 0 || result.ImportAmbiguous.Length != 0 ||
            result.TemporaryCopiesPossiblyRemaining.Length != 0 ||
            result.MetadataRestoreAmbiguous.Length != 0 ||
            result.Items.Any(x => !x.DestinationComparison.EvidenceAvailable || !x.SourceAfterComparison.EvidenceAvailable ||
                x.DestinationComparison.DifferentFields.Length != 0 || x.SourceAfterComparison.DifferentFields.Length != 0);
        return result;
    }

    private static Comparison Compare(string before, string after)
    {
        if (!File.Exists(before) || !File.Exists(after)) return new(false, []);
        var left = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(before))
            ?? throw new InvalidDataException("Hash de campos inválido.");
        var right = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(after))
            ?? throw new InvalidDataException("Hash de campos inválido.");
        if (left.Count == 0 || right.Count == 0) return new(false, []);
        return new(true, left.Keys.Union(right.Keys).Where(key => left.GetValueOrDefault(key) != right.GetValueOrDefault(key)).Order(StringComparer.Ordinal).ToArray());
    }

    private static List<JsonElement> ReadJournal(string path, List<string> errors)
    {
        var events = new List<JsonElement>();
        if (!File.Exists(path)) { errors.Add("Diário ausente: " + path); return events; }
        int line = 0;
        foreach (string text in File.ReadLines(path))
        {
            line++;
            try
            {
                using var json = JsonDocument.Parse(text);
                var entry = json.RootElement.GetProperty("Event");
                if (entry.ValueKind != JsonValueKind.Object) throw new InvalidDataException("Evento não é objeto.");
                events.Add(entry.Clone());
            }
            catch (Exception ex) when (ex is JsonException or KeyNotFoundException or InvalidOperationException or InvalidDataException)
            { errors.Add($"Diário inválido/truncado: {Path.GetFileName(path)}, linha {line}."); }
        }
        return events;
    }
    private static string? State(JsonElement entry) => entry.TryGetProperty("State", out var state) ? state.GetString() : null;
    private static HashSet<int> Indices(List<JsonElement> events, string state) => events.Where(x => State(x) == state)
        .Where(x => x.TryGetProperty("Index", out var value) && value.TryGetInt32(out _)).Select(x => x.GetProperty("Index").GetInt32()).ToHashSet();

    internal static void SelfTest()
    {
        string run = Path.Combine(Path.GetTempPath(), "outlook-audit-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(run);
        try
        {
            File.WriteAllText(Path.Combine(run, "report.json"), JsonSerializer.Serialize(new PilotReport { Status = "review-required" }));
            using (var journal = new DurableJournal(Path.Combine(run, "journal.jsonl")))
            {
                journal.Write(new { State = "source-copy-started", Index = 1 });
                journal.Write(new { State = "source-copy-started", Index = 2 });
                journal.Write(new { State = "source-copy-created", Index = 2 });
                journal.Write(new { State = "import-started", Index = 2 });
                journal.Write(new { State = "unsent-delivery-time-restore-started", Index = 3 });
            }
            var audit = Read(run);
            if (!audit.ReviewRequired || !audit.CopyCreationAmbiguous.SequenceEqual(new[] { 1 }) ||
                !audit.ImportAmbiguous.SequenceEqual(new[] { 2 }) || !audit.TemporaryCopiesPossiblyRemaining.SequenceEqual(new[] { 2 }) ||
                !audit.MetadataRestoreAmbiguous.SequenceEqual(new[] { 3 }))
                throw new Exception("Auditoria ignorou operação ambígua.");
            File.Delete(Path.Combine(run, "report.json"));
            var interrupted = Read(run);
            if (!interrupted.ReviewRequired || !interrupted.ImportAmbiguous.SequenceEqual(new[] { 2 }) || interrupted.Errors.Count == 0)
                throw new Exception("Interrupção sem relatório não auditada.");
            File.WriteAllText(Path.Combine(run, "report.json"), JsonSerializer.Serialize(new PilotReport { Status = "review-required" }));
            File.AppendAllText(Path.Combine(run, "journal.jsonl"), "{\"Event\":");
            if (Read(run).Errors.Count != 1) throw new Exception("Diário truncado não detectado.");
            string left = Path.Combine(run, "left.json"), right = Path.Combine(run, "right.json");
            if (Compare(left, right).EvidenceAvailable) throw new Exception("Evidência ausente aceita.");
            File.WriteAllText(left, "{\"Body\":\"AAA\"}"); File.WriteAllText(right, "{\"Body\":\"BBB\"}");
            if (!Compare(left, right).DifferentFields.SequenceEqual(new[] { "Body" })) throw new Exception("Divergência não detectada.");
            Console.WriteLine("Auditoria: cópia/importação ambíguas, diário truncado e evidência ausente/divergente detectados.");
        }
        finally
        {
            foreach (string name in new[] { "report.json", "journal.jsonl", "left.json", "right.json" }) File.Delete(Path.Combine(run, name));
            Directory.Delete(run);
        }
    }
}
internal sealed class AuditResult
{
    public string RunDirectory { get; set; } = "";
    public string RecordedStatus { get; set; } = "";
    public int RecordedSourceCountBefore { get; set; }
    public int? RecordedSourceCountAfter { get; set; }
    public bool ReviewRequired { get; set; } = true;
    public bool AutomaticRetryAllowed { get; } = false;
    public bool RemovalAllowed { get; } = false;
    public string Scope { get; } = "Evidências locais históricas; não consulta a existência atual dos itens no Outlook.";
    public int[] CopyCreationAmbiguous { get; set; } = [];
    public int[] ImportAmbiguous { get; set; } = [];
    public int[] TemporaryCopiesPossiblyRemaining { get; set; } = [];
    public int[] MetadataRestoreAmbiguous { get; set; } = [];
    public string[] JournalImportedDestinationIds { get; set; } = [];
    public List<AuditItem> Items { get; set; } = new();
    public List<DestinationRelocation> DestinationRelocations { get; set; } = new();
    public List<string> Errors { get; set; } = new();
}
internal sealed class AuditItem
{
    public int Index { get; set; }
    public string RecordedStatus { get; set; } = "";
    public string SourceId { get; set; } = "";
    public string? RecordedDestinationId { get; set; }
    public string? TemporarySourceCopyId { get; set; }
    public Comparison DestinationComparison { get; set; } = new(false, []);
    public Comparison SourceAfterComparison { get; set; } = new(false, []);
}
internal record Comparison(bool EvidenceAvailable, string[] DifferentFields);
internal record DestinationRelocation(string EvidenceDirectory, string PreviousId, string DestinationId);
