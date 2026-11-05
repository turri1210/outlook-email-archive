using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;

// The database transaction covers each state/event pair, never a COM action.
// ImportStarted/DeleteStarted survive a crash and must be reconciled by engine.
internal sealed partial class ArchiveLedger : IDisposable
{
    private readonly SqliteConnection connection;
    internal ArchiveLedger(string path, bool readOnly = false)
    {
        bool existed = File.Exists(path);
        connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = path, Pooling = false,
            Mode = readOnly ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate }.ToString());
        connection.Open();
        if (existed)
        {
            using var check = connection.CreateCommand(); check.CommandText = "PRAGMA application_id";
            long applicationId = Convert.ToInt64(check.ExecuteScalar());
            if (applicationId is not (1128355889 or 1128355890 or 1128355891 or 1128355892 or 1128355893)) throw new InvalidDataException("O arquivo não é um catálogo reconhecido desta ferramenta.");
        }
        if (readOnly) return;
        Execute("PRAGMA journal_mode=WAL; PRAGMA synchronous=FULL; PRAGMA foreign_keys=ON; PRAGMA busy_timeout=5000;");
        Execute("""
            CREATE TABLE IF NOT EXISTS jobs(id TEXT PRIMARY KEY,payload TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS volumes(job TEXT NOT NULL,number INTEGER NOT NULL,version INTEGER NOT NULL,payload TEXT NOT NULL,
              PRIMARY KEY(job,number),FOREIGN KEY(job) REFERENCES jobs(id));
            CREATE TABLE IF NOT EXISTS items(key TEXT PRIMARY KEY,job TEXT NOT NULL,token TEXT NOT NULL UNIQUE,version INTEGER NOT NULL,
              state INTEGER NOT NULL,volume INTEGER,payload TEXT NOT NULL,FOREIGN KEY(job) REFERENCES jobs(id),
              FOREIGN KEY(job,volume) REFERENCES volumes(job,number));
            CREATE INDEX IF NOT EXISTS items_job ON items(job);
            CREATE INDEX IF NOT EXISTS items_job_volume ON items(job,volume);
            CREATE TABLE IF NOT EXISTS batches(id TEXT PRIMARY KEY,job TEXT NOT NULL,version INTEGER NOT NULL,payload TEXT NOT NULL,
              FOREIGN KEY(job) REFERENCES jobs(id));
            CREATE TABLE IF NOT EXISTS events(id INTEGER PRIMARY KEY AUTOINCREMENT,job TEXT NOT NULL,utc TEXT NOT NULL,kind TEXT NOT NULL,payload TEXT NOT NULL,
              FOREIGN KEY(job) REFERENCES jobs(id));
            """);
        if (!existed) Execute("PRAGMA application_id=1128355889; PRAGMA user_version=1;");
        InitializeWorkflow();
        InitializeHistoryStorage();
        InitializeEvidenceMaintenance();
        InitializePlanning();
    }
    internal void InsertJob(ArchiveJob job)
    {
        if (!Guid.TryParseExact(job.Id, "N", out _)) throw new ArgumentException("Job deve ser GUID N.");
        using var config = JsonDocument.Parse(job.ConfigJson);
        if (config.RootElement.ValueKind != JsonValueKind.Object) throw new ArgumentException("Configuração deve ser objeto JSON.");
        using var transaction = connection.BeginTransaction();
        Run("INSERT INTO jobs(id,payload) VALUES($id,$payload)", transaction, ("$id", job.Id), ("$payload", JsonSerializer.Serialize(job)));
        Event(job.Id, "job-created", job, transaction); transaction.Commit();
    }
    internal ArchiveJob GetJob(string id) => Single<ArchiveJob>("SELECT payload FROM jobs WHERE id=$id", ("$id", id)) ?? throw new KeyNotFoundException("Execução não encontrada.");
    internal List<ArchiveJob> GetJobs() => Many<ArchiveJob>("SELECT payload FROM jobs ORDER BY rowid");
    internal HashSet<string> GetKnownItemKeys()
    {
        using var command = connection.CreateCommand(); command.CommandText = "SELECT key FROM items";
        using var reader = command.ExecuteReader(); var result = new HashSet<string>(StringComparer.Ordinal);
        while (reader.Read()) result.Add(reader.GetString(0));
        reader.Close();
        using var moved = connection.CreateCommand();
        moved.CommandText = "SELECT json_extract(payload,'$.SourceStoreId'),json_extract(payload,'$.RemovalEntryId') FROM items WHERE json_extract(payload,'$.RemovalEntryId') IS NOT NULL";
        using var rows = moved.ExecuteReader();
        while (rows.Read()) result.Add(ComputeItemKey(rows.GetString(0), rows.GetString(1)));
        return result;
    }
    internal bool HasUnreconciledRemoval(string storeId)
    {
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM items WHERE state NOT IN ($removed,$confirmed) AND json_extract(payload,'$.SourceStoreId')=$store AND json_extract(payload,'$.RemovalProofPath') IS NOT NULL AND json_extract(payload,'$.RemovalEntryId') IS NULL";
        command.Parameters.AddWithValue("$removed", (int)ArchiveItemState.Removed);
        command.Parameters.AddWithValue("$confirmed", (int)ArchiveItemState.BackupVerifiedSourceUnmanaged); command.Parameters.AddWithValue("$store", storeId);
        return Convert.ToInt64(command.ExecuteScalar()) != 0;
    }
    internal static string ComputeItemKey(string storeId, string entryId)
    {
        if (string.IsNullOrWhiteSpace(storeId) || string.IsNullOrWhiteSpace(entryId)) throw new ArgumentException("Identidades vazias.");
        byte[] a = Encoding.UTF8.GetBytes(storeId), b = Encoding.UTF8.GetBytes(entryId);
        using var stream = new MemoryStream(); using var writer = new BinaryWriter(stream);
        writer.Write(a.Length); writer.Write(a); writer.Write(b.Length); writer.Write(b);
        return Convert.ToHexString(SHA256.HashData(stream.ToArray()));
    }
    internal bool AddItem(ArchiveItem item)
    {
        string key = ComputeItemKey(item.SourceStoreId, item.SourceEntryId);
        if (item.Key.Length == 0) item.Key = key;
        if (item.Key != key || item.Version != 0 || !Guid.TryParseExact(item.Token, "N", out _)) throw new ArgumentException("Identidade/token/versão inválidos.");
        using var tx = connection.BeginTransaction();
        int inserted = Run("INSERT INTO items(key,job,token,version,state,volume,payload) VALUES($key,$job,$token,0,$state,$volume,$payload) ON CONFLICT(key) DO NOTHING", tx,
            ("$key", item.Key), ("$job", item.JobId), ("$token", item.Token), ("$state", (int)item.State),
            ("$volume", item.VolumeNumber), ("$payload", JsonSerializer.Serialize(item)));
        if (inserted == 0) return false;
        Event(item.JobId, "item-added", item, tx); tx.Commit(); return true;
    }
    internal ArchiveItem? GetItem(string key) => Single<ArchiveItem>("SELECT payload FROM items WHERE key=$key", ("$key", key));
    internal List<ArchiveItem> GetItems(string job) => Many<ArchiveItem>("SELECT payload FROM items WHERE job=$job ORDER BY rowid", ("$job", job));
    internal List<ArchiveBatch> GetBatches(string job)
    {
        using var exists = connection.CreateCommand(); exists.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='batches'";
        if (Convert.ToInt64(exists.ExecuteScalar()) == 0) return [];
        return Many<ArchiveBatch>("SELECT payload FROM batches WHERE job=$job ORDER BY rowid", ("$job", job));
    }
    internal void AddBatch(ArchiveBatch batch, IReadOnlyList<ArchiveItem> items)
    {
        if (!Guid.TryParseExact(batch.Id, "N", out _) || batch.Version != 0 || batch.State != ArchiveBatchState.Prepared ||
            (items.Count < 1 || items.Count > (batch.FastCopy ? 100 : 50)) || batch.Items.Count != items.Count || batch.Items.Distinct().Count() != items.Count ||
            !batch.Items.SequenceEqual(items.Select(x => x.Key)) || items.Any(x => x.JobId != batch.JobId || x.NativeBatchId is not null || (batch.FastCopy ? x.State != ArchiveItemState.Planned || x.FastIdentityHash is null || x.FastIdentityPath is null : x.State != ArchiveItemState.Exported)))
            throw new InvalidOperationException("Membros/estado do lote inválidos.");
        using var tx = connection.BeginTransaction();
        // Old binaries reject this application_id before opening a write session.
        // The batch and every member link are committed in the same transaction.
        using (var schema = connection.CreateCommand())
        {
            schema.Transaction = tx; schema.CommandText = "PRAGMA application_id";
            if (Convert.ToInt64(schema.ExecuteScalar()) is not (1128355891 or 1128355892 or 1128355893)) Run("PRAGMA application_id=1128355890; PRAGMA user_version=2;", tx);
        }
        Run("INSERT INTO batches(id,job,version,payload) VALUES($id,$job,0,$payload)", tx,
            ("$id", batch.Id), ("$job", batch.JobId), ("$payload", JsonSerializer.Serialize(batch)));
        foreach (var row in items)
        {
            var changed = JsonSerializer.Deserialize<ArchiveItem>(JsonSerializer.Serialize(row))!;
            changed.NativeBatchId = batch.Id; changed.Version++;
            if (Run("UPDATE items SET version=$next,payload=$payload WHERE key=$key AND version=$previous", tx,
                ("$next", changed.Version), ("$payload", JsonSerializer.Serialize(changed)), ("$key", row.Key), ("$previous", row.Version)) != 1)
                throw new InvalidOperationException("Conflito ao vincular o lote.");
            Event(row.JobId, "batch-item-linked", changed, tx);
        }
        Event(batch.JobId, "batch-created", batch, tx); tx.Commit();
        foreach (var row in items) { row.NativeBatchId = batch.Id; row.Version++; }
    }
    internal void UpdateBatch(ArchiveBatch batch)
    {
        var old = GetBatches(batch.JobId).Single(x => x.Id == batch.Id);
        if (old.Version != batch.Version || old.ManifestSha256 != batch.ManifestSha256 || old.DirectoryPath != batch.DirectoryPath ||
            old.VolumeNumber != batch.VolumeNumber || old.SourceStoreId != batch.SourceStoreId || old.SourceFolderId != batch.SourceFolderId ||
            !old.Items.SequenceEqual(batch.Items) || !old.RelativePath.SequenceEqual(batch.RelativePath) ||
            old.State == ArchiveBatchState.Complete && batch.State != ArchiveBatchState.Complete)
            throw new InvalidOperationException("Identidade/versão do lote inválida.");
        var next = JsonSerializer.Deserialize<ArchiveBatch>(JsonSerializer.Serialize(batch))!; next.Version++;
        using var tx = connection.BeginTransaction();
        if (Run("UPDATE batches SET version=$next,payload=$payload WHERE id=$id AND version=$old", tx,
            ("$next", next.Version), ("$payload", JsonSerializer.Serialize(next)), ("$id", batch.Id), ("$old", batch.Version)) != 1)
            throw new InvalidOperationException("Conflito de versão do lote.");
        Event(batch.JobId, "batch-updated", next, tx); tx.Commit(); batch.Version = next.Version;
    }
    internal long GetReservedBytes(string job, int volume, string excludedKey)
    {
        // Aggregate only the requested part in SQLite. Do not materialize every
        // message payload merely to budget the next import. Missing fields in
        // older payloads have the same zero default as ArchiveItem.ReservedBytes.
        using var command = Command("""
            SELECT COALESCE(SUM(COALESCE(json_extract(payload,'$.ReservedBytes'),0)),0)
            FROM items WHERE job=$job AND volume=$volume AND key<>$excluded
            """, null, ("$job", job), ("$volume", volume), ("$excluded", excludedKey));
        return Convert.ToInt64(command.ExecuteScalar());
    }
    internal void UpdateItem(ArchiveItem item)
    {
        var old = GetItem(item.Key) ?? throw new KeyNotFoundException("Item não registrado.");
        if (old.JobId != item.JobId || old.Token != item.Token || old.SourceStoreId != item.SourceStoreId || old.SourceEntryId != item.SourceEntryId ||
            old.NativeBatchId != item.NativeBatchId ||
            (old.RepairToken is not null && (old.RepairToken != item.RepairToken || old.RepairManifestSha256 != item.RepairManifestSha256 || old.RepairManifestPath != item.RepairManifestPath)) ||
            (old.State == ArchiveItemState.BackupVerifiedSourceUnmanaged && item.State != old.State) ||
            (old.State == ArchiveItemState.SkippedByPolicy && item.State != old.State) ||
            (old.State == ArchiveItemState.Removed && item.State != ArchiveItemState.Removed)) throw new InvalidOperationException("Identidade/estado final imutável.");
        int previous = item.Version;
        using var tx = connection.BeginTransaction();
        if (item.RepairToken is not null || item.State == ArchiveItemState.BackupVerifiedSourceUnmanaged)
        {
            using var schema=Command("PRAGMA application_id",tx);
            if(Convert.ToInt64(schema.ExecuteScalar()) is not (1128355892 or 1128355893))Run("PRAGMA application_id=1128355891; PRAGMA user_version=3;",tx);
        }
        var changed = JsonSerializer.Deserialize<ArchiveItem>(JsonSerializer.Serialize(item))!; changed.Version++;
        int rows = Run("UPDATE items SET version=$next,state=$state,volume=$volume,payload=$payload WHERE key=$key AND version=$previous", tx,
            ("$next", changed.Version), ("$state", (int)changed.State), ("$volume", changed.VolumeNumber), ("$payload", JsonSerializer.Serialize(changed)),
            ("$key", item.Key), ("$previous", previous));
        if (rows != 1) throw new InvalidOperationException("Conflito de versão; operação interrompida.");
        Run("UPDATE workflow_batches SET state='Pending' WHERE job=$job AND state IN ('Complete','Review') AND number IN (SELECT batch FROM workflow_members WHERE key=$key)",tx,("$job",item.JobId),("$key",item.Key));
        Event(item.JobId, "item-updated", changed, tx); tx.Commit(); item.Version = changed.Version;
    }
    internal bool AddVolume(ArchiveVolume volume)
    {
        if (volume.Number < 1 || volume.Version != 0 || !Path.IsPathFullyQualified(volume.PstPath)) throw new ArgumentException("Volume inválido.");
        using var tx = connection.BeginTransaction();
        int inserted = Run("INSERT INTO volumes(job,number,version,payload) VALUES($job,$number,0,$payload) ON CONFLICT(job,number) DO NOTHING", tx,
            ("$job", volume.JobId), ("$number", volume.Number), ("$payload", JsonSerializer.Serialize(volume)));
        if (inserted == 0) return false;
        Event(volume.JobId, "volume-added", volume, tx); tx.Commit(); return true;
    }
    internal List<ArchiveVolume> GetVolumes(string job) => Many<ArchiveVolume>("SELECT payload FROM volumes WHERE job=$job ORDER BY number", ("$job", job));
    internal void UpdateVolume(ArchiveVolume volume)
    {
        var old = GetVolumes(volume.JobId).Single(x => x.Number == volume.Number);
        if (!string.Equals(old.PstPath, volume.PstPath, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("Caminho de volume imutável.");
        if (old.IsExisting != volume.IsExisting || old.SourceStoreId != volume.SourceStoreId) throw new InvalidOperationException("Identidade do volume imutável.");
        if (old.PreserveFolderStructure != volume.PreserveFolderStructure || JsonSerializer.Serialize(old.FolderMappings) != JsonSerializer.Serialize(volume.FolderMappings) ||
            (old.ControlFolderId is not null && old.ControlFolderId != volume.ControlFolderId)) throw new InvalidOperationException("Mapeamento de pastas do volume imutável.");
        if (old.Version != volume.Version) throw new InvalidOperationException("Conflito de versão do volume.");
        if (old.StoreId == volume.StoreId && old.State == volume.State && old.ObservedBytes == volume.ObservedBytes && old.InitialBytes == volume.InitialBytes && old.ControlFolderId == volume.ControlFolderId)
            return; // No state changed: avoid a redundant FULL-sync transaction/event.
        var changed = JsonSerializer.Deserialize<ArchiveVolume>(JsonSerializer.Serialize(volume))!; changed.Version++;
        using var tx = connection.BeginTransaction();
        if (Run("UPDATE volumes SET version=$next,payload=$payload WHERE job=$job AND number=$number AND version=$old", tx,
            ("$next", changed.Version), ("$payload", JsonSerializer.Serialize(changed)), ("$job", changed.JobId), ("$number", changed.Number), ("$old", volume.Version)) != 1)
            throw new InvalidOperationException("Conflito de versão do volume.");
        Event(volume.JobId, "volume-updated", changed, tx); tx.Commit(); volume.Version = changed.Version;
    }
    internal List<string> GetEvents(string job)
    {
        return ReadHistory(job).OrderBy(x=>x.Id).Select(x=>x.Payload).ToList();
    }
    private void Event<T>(string job, string kind, T payload, SqliteTransaction tx) => Run(
        "INSERT INTO events(job,utc,kind,payload) VALUES($job,$utc,$kind,$payload)", tx,
        ("$job", job), ("$utc", DateTimeOffset.UtcNow.ToString("O")), ("$kind", kind), ("$payload", JsonSerializer.Serialize(payload)));
    private T? Single<T>(string sql, params (string, object?)[] values) where T : class => Many<T>(sql, values).SingleOrDefault();
    private List<T> Many<T>(string sql, params (string, object?)[] values)
    {
        using var cmd = Command(sql, null, values); using var reader = cmd.ExecuteReader(); var result = new List<T>();
        while (reader.Read()) result.Add(JsonSerializer.Deserialize<T>(reader.GetString(0)) ?? throw new InvalidDataException("Estado SQLite inválido."));
        return result;
    }
    private int Run(string sql, SqliteTransaction? tx, params (string, object?)[] values)
    { using var cmd = Command(sql, tx, values); return cmd.ExecuteNonQuery(); }
    private SqliteCommand Command(string sql, SqliteTransaction? tx, params (string, object?)[] values)
    {
        var cmd = connection.CreateCommand(); cmd.CommandText = sql; cmd.Transaction = tx;
        foreach (var (name, value) in values) cmd.Parameters.AddWithValue(name, value ?? DBNull.Value); return cmd;
    }
    private void Execute(string sql) => Run(sql, null);
    public void Dispose() => connection.Dispose();
}
