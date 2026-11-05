using System.Text.Json;

// Bounded copy-only release check. The user's catalog and destination PST are
// never opened for writing; only a new private diagnostic catalog is created.
internal static class NativeBatchSmoke
{
    internal static int Execute(string[] args, bool fast = false)
    {
        if (args.Length != 3) throw new ArgumentException("Informe raiz, execução e nova pasta de teste.");
        int sampleCount = fast && new[] { "outlook-fast-integration-m365-20261027-06", "outlook-fast-integration-m365-20261027-07" }.Contains(Path.GetFileName(args[2])) ? 150 : fast ? 50 : 5;
        var elapsed = System.Diagnostics.Stopwatch.StartNew();
        string sourceRoot = Path.GetFullPath(args[0]), targetRoot = Path.GetFullPath(args[2]);
        ArchiveFiles.ValidateLocalDirectory(targetRoot);
        string downloads = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        if (!PilotCommand.IsWithin(targetRoot, downloads) || (fast ? !new[] { "outlook-fast-integration-m365-20261027-01", "outlook-fast-integration-m365-20261027-02", "outlook-fast-integration-m365-20261027-03", "outlook-fast-integration-m365-20261027-04", "outlook-fast-integration-m365-20261027-05", "outlook-fast-integration-m365-20261027-06", "outlook-fast-integration-m365-20261027-07" }.Contains(Path.GetFileName(targetRoot)) : Path.GetFileName(targetRoot) != "outlook-native-integration-m365-20261027-01") || Directory.Exists(targetRoot))
            throw new ArgumentException("O ensaio exige uma pasta privada nova e identificada em Downloads.");
        using var sourceLease = new FileStream(Path.Combine(sourceRoot, "archive.lock"), FileMode.Open, FileAccess.Read, FileShare.None);
        using var read = new ArchiveLedger(Path.Combine(sourceRoot, "archive.sqlite"), true);
        var previous = read.GetJob(args[1]);
        var envelope = JsonSerializer.Deserialize<ArchivePlanEnvelope>(previous.ConfigJson)!;
        string sourcePlan = Path.Combine(previous.RootDirectory, "source-plan.json");
        if (ArchiveFiles.Sha(sourcePlan) != envelope.PlanSha256) throw new ArchiveReviewException("Plano original alterado.");
        var descriptors = JsonSerializer.Deserialize<List<SourceDescriptor>>(File.ReadAllText(sourcePlan))!
            .ToDictionary(x => ArchiveLedger.ComputeItemKey(x.StoreId, x.EntryId), StringComparer.Ordinal);
        using var backend = new OutlookArchiveBackend();
        var pool = read.GetItems(previous.Id).Where(x => x.State == ArchiveItemState.Planned && x.SourceFolderId == envelope.Config.SourceFolderId).Take(300).ToList();
        if (fast) backend.RefreshSourceIdentities(pool);
        var selected = pool.Where(row => { try { return backend.CanBatch(row); } catch (ArchiveReviewException) { return false; } }).Take(sampleCount).ToList();
        backend.ClearFastIdentityCache();
        if (selected.Count != sampleCount || selected.Any(x => !descriptors.ContainsKey(x.Key))) throw new ArchiveReviewException("Amostra de itens congelados indisponível.");
        var config = JsonSerializer.Deserialize<ArchiveConfig>(JsonSerializer.Serialize(envelope.Config))!;
        config.RootDirectory = targetRoot; config.ExistingPstPath = null; config.IncludeSubfolders = false;
        config.AllowRetainedMsgBackupForRemoval = true;
        Directory.CreateDirectory(targetRoot);
        string id = Guid.NewGuid().ToString("N"), jobDirectory = Path.Combine(targetRoot, "jobs", id);
        Directory.CreateDirectory(Path.Combine(jobDirectory, "psts"));
        var rows = selected.Select(x => descriptors[x.Key]).ToList();
        string plan = Path.Combine(jobDirectory, "source-plan.json"), scopePath = Path.Combine(jobDirectory, "source-scope.json");
        ArchiveFiles.WriteNew(plan, rows);
        ArchiveFiles.WriteNew(scopePath, new[] { new FolderScopeSnapshot(config.SourceFolderId, rows[0].RelativePath, rows.Select(x => x.EntryId).ToList()) });
        using (var ledger = new ArchiveLedger(Path.Combine(targetRoot, "archive.sqlite")))
        {
            ledger.InsertJob(new ArchiveJob { Id = id, RootDirectory = jobDirectory,
                ConfigJson = JsonSerializer.Serialize(new ArchivePlanEnvelope(config, ArchiveFiles.Sha(plan), ArchiveFiles.Sha(scopePath))) });
        }
        using var engine = new ArchiveEngine(targetRoot, backend);
        var copyElapsed = System.Diagnostics.Stopwatch.StartNew();
        engine.Copy(id); copyElapsed.Stop();
        if (!fast) engine.Verify(id);
        var items = engine.Items(id);
        var result = new { CopyOnly = true, MainCatalogReadOnly = true, RemovalAllowed = false, JobId = id,
            FastCopy = fast, CopySeconds = copyElapsed.Elapsed.TotalSeconds, TotalSeconds = elapsed.Elapsed.TotalSeconds,
            Copied = items.Count(x => x.State == ArchiveItemState.Copied), Count = items.Count, Verified = items.Count(x => x.State == ArchiveItemState.Verified),
            NativeItems = items.Count(x => x.NativeBatchId is not null), Errors = items.Where(x => x.Error is not null && !(fast && x.State == ArchiveItemState.Copied && x.Error.StartsWith("Cópia rápida concluída"))).Select(x => new { x.Key, x.Error }).ToArray(),
            PstPaths = engine.Volumes(id).Select(x => x.PstPath).ToArray() };
        ArchiveFiles.WriteNew(Path.Combine(targetRoot, "release-smoke-summary.json"), result);
        Console.WriteLine(JsonSerializer.Serialize(result, ArchiveFiles.JsonOptions));
        return result.Count == sampleCount && (fast ? result.Copied == sampleCount : result.Verified == sampleCount) && result.NativeItems == sampleCount && result.Errors.Length == 0 ? 0 : 3;
    }
}
