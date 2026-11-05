using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;

internal sealed partial class ArchiveLedger
{
    private sealed record HistoryRow(long Id,string Job,string Utc,string Kind,string Payload);
    private void InitializeHistoryStorage()=>Execute("""
        CREATE INDEX IF NOT EXISTS events_job_id ON events(job,id);
        CREATE TABLE IF NOT EXISTS event_archives(first_id INTEGER PRIMARY KEY,last_id INTEGER NOT NULL,job TEXT NOT NULL,count INTEGER NOT NULL,sha TEXT NOT NULL,raw_bytes INTEGER NOT NULL,data BLOB NOT NULL,FOREIGN KEY(job) REFERENCES jobs(id));
        CREATE INDEX IF NOT EXISTS event_archive_job ON event_archives(job,first_id);
        """);
    private static byte[] InflateHistory(byte[] data,int size)
    {
        if(size<1 || size>64*1024*1024)throw new InvalidDataException("Tamanho do histórico comprimido inválido.");
        using var source=new MemoryStream(data,false);using var gzip=new GZipStream(source,CompressionMode.Decompress);
        var bytes=new byte[size];gzip.ReadExactly(bytes);if(gzip.ReadByte()!=-1)throw new InvalidDataException("Histórico descomprimido excede seu tamanho registrado.");return bytes;
    }
    private List<HistoryRow> ReadHistory(string job)
    {
        var result=new List<HistoryRow>();
        using(var command=Command("SELECT id,job,utc,kind,payload FROM events WHERE job=$job ORDER BY id",null,("$job",job)))
        using(var reader=command.ExecuteReader())while(reader.Read())result.Add(new(reader.GetInt64(0),reader.GetString(1),reader.GetString(2),reader.GetString(3),reader.GetString(4)));
        using(var exists=Command("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='event_archives'",null))if(Convert.ToInt64(exists.ExecuteScalar())==0)return result;
        using var cmd=Command("SELECT first_id,last_id,count,sha,raw_bytes,data FROM event_archives WHERE job=$job ORDER BY first_id",null,("$job",job));using var cursor=cmd.ExecuteReader();
        while(cursor.Read())
        {
            byte[] bytes=InflateHistory((byte[])cursor[5],cursor.GetInt32(4));
            if(Convert.ToHexString(SHA256.HashData(bytes))!=cursor.GetString(3))throw new InvalidDataException("Integridade do histórico comprimido não confere.");
            var rows=JsonSerializer.Deserialize<List<HistoryRow>>(bytes)??throw new InvalidDataException("Histórico comprimido inválido.");
            if(rows.Count!=cursor.GetInt32(2)||rows[0].Id!=cursor.GetInt64(0)||rows[^1].Id!=cursor.GetInt64(1)||rows.Any(x=>x.Job!=job)||rows.Select(x=>x.Id).Distinct().Count()!=rows.Count)
                throw new InvalidDataException("Índice do histórico comprimido não confere.");
            result.AddRange(rows);
        }
        if(result.Select(x=>x.Id).Distinct().Count()!=result.Count)throw new InvalidDataException("Histórico contém eventos duplicados.");return result;
    }
    internal int PackHistory(string job,int keepRecent=2000)
    {
        if(keepRecent<0)throw new ArgumentOutOfRangeException(nameof(keepRecent));
        long cutoff;
        using(var command=Command("SELECT id FROM events WHERE job=$job ORDER BY id DESC LIMIT 1 OFFSET $keep",null,("$job",job),("$keep",keepRecent)))
        {var value=command.ExecuteScalar();if(value is null)return 0;cutoff=Convert.ToInt64(value);}
        int packed=0;
        while(true)
        {
            var rows=new List<HistoryRow>();long characters=0;
            using(var cmd=Command("SELECT id,job,utc,kind,payload FROM events WHERE job=$job AND id<=$last ORDER BY id LIMIT 500",null,("$job",job),("$last",cutoff)))
            using(var reader=cmd.ExecuteReader())while(reader.Read())
            {
                string payload=reader.GetString(4);if(rows.Count>0&&characters+payload.Length>4*1024*1024)break;
                rows.Add(new(reader.GetInt64(0),reader.GetString(1),reader.GetString(2),reader.GetString(3),payload));characters+=payload.Length;
            }
            if(rows.Count==0)break;
            byte[] bytes=JsonSerializer.SerializeToUtf8Bytes(rows);if(bytes.Length>64*1024*1024)throw new InvalidDataException("Evento de histórico excede o limite de compactação.");
            string sha=Convert.ToHexString(SHA256.HashData(bytes));byte[] compressed;
            using(var target=new MemoryStream()){using(var gzip=new GZipStream(target,CompressionLevel.Fastest,true))gzip.Write(bytes);compressed=target.ToArray();}
            if(!InflateHistory(compressed,bytes.Length).AsSpan().SequenceEqual(bytes))throw new InvalidDataException("Verificação do histórico comprimido falhou.");
            using var tx=connection.BeginTransaction();
            Run("PRAGMA application_id=1128355893; PRAGMA user_version=5;",tx);
            Run("INSERT INTO event_archives(first_id,last_id,job,count,sha,raw_bytes,data) VALUES($f,$l,$j,$c,$s,$r,$d)",tx,("$f",rows[0].Id),("$l",rows[^1].Id),("$j",job),("$c",rows.Count),("$s",sha),("$r",bytes.Length),("$d",compressed));
            int removed=Run("DELETE FROM events WHERE job=$job AND id>=$first AND id<=$last",tx,("$job",job),("$first",rows[0].Id),("$last",rows[^1].Id));
            if(removed!=rows.Count)throw new InvalidOperationException("O histórico mudou durante a compactação.");
            tx.Commit();packed+=rows.Count;
        }
        return packed;
    }
    internal void MaybePackHistory(string job)
    {
        using var count=Command("SELECT COUNT(*) FROM events WHERE job=$j",null,("$j",job));
        if(Convert.ToInt64(count.ExecuteScalar())>5000)PackHistory(job);
    }
    internal void ReclaimHistorySpace()
    {
        Execute("PRAGMA wal_checkpoint(TRUNCATE); VACUUM; PRAGMA wal_checkpoint(TRUNCATE);");
    }
}
