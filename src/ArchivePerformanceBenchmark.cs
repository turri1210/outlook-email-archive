using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

// Isolated aggregation microbenchmark. It never opens Outlook or an existing
// user catalog and does not change production checkpoint/durability settings.
internal static class ArchivePerformanceBenchmark
{
    internal static int Run()
    {
        const int count = 6107, iterations = 5;
        string directory = Path.Combine(Path.GetTempPath(), "outlook-ledger-perf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "synthetic.sqlite");
        using var ledger = new ArchiveLedger(path);
        string job = Guid.NewGuid().ToString("N");
        ledger.InsertJob(new() { Id = job, RootDirectory = directory, ConfigJson = "{}" });
        ledger.AddVolume(new() { JobId = job, Number = 1, PstPath = Path.Combine(directory, "synthetic-part-0001.pst"), State = ArchiveVolumeState.Ready });
        long totalPayloadBytes = 0;
        string excludedKey = "";
        using (var connection = new SqliteConnection(new SqliteConnectionStringBuilder
        { DataSource = path, Mode = SqliteOpenMode.ReadWrite, Pooling = false }.ToString()))
        {
            connection.Open();
            using (var settings = connection.CreateCommand())
            {
                settings.CommandText = "PRAGMA synchronous=FULL; PRAGMA foreign_keys=ON; PRAGMA busy_timeout=5000;";
                settings.ExecuteNonQuery();
            }
            // Fixture preparation is deliberately one batch. Timings below
            // measure the read paths, never synthetic per-item insert latency.
            using var transaction = connection.BeginTransaction();
            using var insert = connection.CreateCommand(); insert.Transaction = transaction;
            insert.CommandText = "INSERT INTO items(key,job,token,version,state,volume,payload) VALUES($key,$job,$token,0,$state,1,$payload)";
            insert.Parameters.Add("$key", SqliteType.Text); insert.Parameters.AddWithValue("$job", job);
            insert.Parameters.Add("$token", SqliteType.Text); insert.Parameters.AddWithValue("$state", (int)ArchiveItemState.Copied);
            insert.Parameters.Add("$payload", SqliteType.Text); insert.Prepare();
            for (int i = 0; i < count; i++)
            {
                string id = "synthetic-entry-" + i.ToString("D8");
                string key = ArchiveLedger.ComputeItemKey("synthetic-source-store", id);
                string evidence = Path.Combine(directory, "jobs", job, "i", key[..16], "original-synthetic-0001");
                var item = new ArchiveItem
                {
                    Key = key, JobId = job, SourceStoreId = "synthetic-source-store", SourceFolderId = "synthetic-source-folder",
                    SourceEntryId = id, SourceRecordKey = new string('A', 64), SourceChangeKey = new string('B', 64),
                    RelativePath = ["Fixture", "Recebidos", "Arquivo de benchmark sintético"], EstimatedBytes = 800 * 1024,
                    State = ArchiveItemState.Copied, VolumeNumber = 1, ReservedBytes = ArchiveVolumePolicy.MiB,
                    SnapshotHash = new string('C', 64), StageMsgPath = Path.Combine(evidence, "message.msg"), StageMsgSha256 = new string('D', 64),
                    TargetStoreId = "synthetic-target-store", TargetFolderId = "synthetic-target-folder", TargetEntryId = "synthetic-target-" + id,
                    TargetSnapshotHash = new string('E', 64), TargetFieldsPath = Path.Combine(evidence, "target-baseline-synthetic", "field-hashes.json")
                };
                string payload = JsonSerializer.Serialize(item); totalPayloadBytes += Encoding.UTF8.GetByteCount(payload);
                insert.Parameters["$key"].Value = key; insert.Parameters["$token"].Value = item.Token; insert.Parameters["$payload"].Value = payload;
                insert.ExecuteNonQuery(); if (i == 0) excludedKey = key;
            }
            transaction.Commit();
        }
        long OldSum() => ledger.GetItems(job).Where(x => x.VolumeNumber == 1 && x.Key != excludedKey).Sum(x => x.ReservedBytes);
        long NewSum() => ledger.GetReservedBytes(job, 1, excludedKey);
        long expected = checked((count - 1L) * ArchiveVolumePolicy.MiB);
        if (OldSum() != expected || NewSum() != expected) throw new InvalidDataException("Agregações sintéticas divergiram no aquecimento.");
        var oldTimes = new List<double>(); var newTimes = new List<double>();
        double Measure(Func<long> action)
        {
            long started = Stopwatch.GetTimestamp(); long actual = action();
            double elapsed = Stopwatch.GetElapsedTime(started).TotalMilliseconds;
            if (actual != expected) throw new InvalidDataException("Agregação sintética mudou durante a medição."); return elapsed;
        }
        for (int i = 0; i < iterations; i++)
        {
            // Alternate order to reduce systematic cache-order bias.
            if (i % 2 == 0) { oldTimes.Add(Measure(OldSum)); newTimes.Add(Measure(NewSum)); }
            else { newTimes.Add(Measure(NewSum)); oldTimes.Add(Measure(OldSum)); }
        }
        double Median(List<double> samples) => samples.OrderBy(x => x).ElementAt(samples.Count / 2);
        double oldMs = Median(oldTimes), newMs = Median(newTimes);
        var report = new
        {
            N = count, Iterations = iterations, AveragePayloadBytes = totalPayloadBytes / count, ReservedBytes = expected,
            OldMs = oldMs, NewMs = newMs, Gain = newMs > 0 ? (double?)(oldMs / newMs) : null,
            OldSamplesMs = oldTimes, NewSamplesMs = newTimes,
            Scope = "Microbenchmark sintético da soma de reservas; não mede Outlook, cópia de mensagens ou checkpoints por item.",
            Setup = "Catálogo temporário novo; 6.107 registros artificiais inseridos em uma transação; WAL e synchronous=FULL preservados.",
            EvidenceDirectory = directory, Utc = DateTimeOffset.UtcNow
        };
        ArchiveFiles.WriteNew(Path.Combine(directory, "benchmark.json"), report);
        Console.WriteLine(JsonSerializer.Serialize(report, ArchiveFiles.JsonOptions)); return 0;
    }
}
