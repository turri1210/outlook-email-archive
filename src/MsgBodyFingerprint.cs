using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OpenMcdf;

// Reads only native root-level MSG body streams. No Outlook/COM rendering,
// transcoding, whitespace normalization, HTML parsing or RTF decompression occurs.
// Equality proves byte preservation between serialized MSG files, not fidelity
// to the Exchange server or completeness of all MAPI message properties.
internal static class MsgBodyFingerprint
{
    private static readonly string[] BodyStreams =
    [
        "__substg1.0_1000001F", // PR_BODY, PT_UNICODE
        "__substg1.0_1000001E", // PR_BODY, PT_STRING8
        "__substg1.0_10130102", // PR_HTML, PT_BINARY
        "__substg1.0_1013001F", // HTML, PT_UNICODE variant
        "__substg1.0_10090102"  // PR_RTF_COMPRESSED, PT_BINARY
    ];

    internal static SortedDictionary<string, string> Read(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var root = RootStorage.OpenRead(path);
        var fields = new SortedDictionary<string, string>(StringComparer.Ordinal);
        bool hasBodyStream = false;
        foreach (string name in BodyStreams)
        {
            string key = "MSGBodyStream:" + name;
            if (!root.TryOpenStream(name, out var stream))
            {
                // Always include absence: losing any representation changes the result.
                fields.Add(key, "missing");
                continue;
            }
            using (stream)
            {
                long length = stream.Length;
                string hash = Convert.ToHexString(SHA256.HashData(stream));
                if (stream.Position != length)
                    throw new InvalidDataException($"Leitura incompleta do corpo MSG: {name}.");
                fields.Add(key, "length=" + length.ToString(CultureInfo.InvariantCulture) + ";sha256=" + hash);
                hasBodyStream = true;
            }
        }
        if (!hasBodyStream)
            throw new InvalidDataException("MSG sem stream de corpo reconhecido na raiz; conferência indisponível.");
        return fields;
    }

    internal static int SelfTest()
    {
        string directory = Path.Combine(Path.GetTempPath(), "outlook-msg-body-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string[] names = ["baseline.msg", "same.msg", "changed.msg", "missing.msg", "empty.msg", "no-body.msg", "nested-only.msg", "invalid.msg"];
        try
        {
            string bodyName = BodyStreams[0], htmlName = BodyStreams[2];
            byte[] body = Encoding.Unicode.GetBytes("Olá, corpo nativo.\r\n");
            byte[] html = Encoding.UTF8.GetBytes("<p>Olá, corpo nativo.</p>");
            string FilePath(string name) => Path.Combine(directory, name);

            WriteFixture(FilePath("baseline.msg"), new() { [bodyName] = body, [htmlName] = html });
            WriteFixture(FilePath("same.msg"), new() { [bodyName] = body, [htmlName] = html });
            WriteFixture(FilePath("changed.msg"), new() { [bodyName] = Encoding.Unicode.GetBytes("Olá, corpo alterado.\r\n"), [htmlName] = html });
            WriteFixture(FilePath("missing.msg"), new() { [bodyName] = body });
            WriteFixture(FilePath("empty.msg"), new() { [bodyName] = body, [htmlName] = [] });
            WriteFixture(FilePath("no-body.msg"), new() { ["unrelated"] = [1, 2, 3] });
            WriteFixture(FilePath("nested-only.msg"), new(), nestedBody: body);
            File.WriteAllText(FilePath("invalid.msg"), "Isto não é CFB.");

            var baseline = Read(FilePath("baseline.msg"));
            if (baseline.Count != BodyStreams.Length || !baseline.SequenceEqual(Read(FilePath("same.msg"))))
                throw new Exception("Corpos MSG idênticos não produziram resultados iguais.");
            string bodyKey = "MSGBodyStream:" + bodyName;
            string expectedBody = "length=" + body.Length.ToString(CultureInfo.InvariantCulture) +
                ";sha256=" + Convert.ToHexString(SHA256.HashData(body));
            if (baseline[bodyKey] != expectedBody)
                throw new Exception("Hash ou tamanho não corresponde aos bytes nativos do corpo.");
            if (baseline.SequenceEqual(Read(FilePath("changed.msg"))))
                throw new Exception("Alteração do corpo não foi detectada.");
            var missing = Read(FilePath("missing.msg"));
            if (missing["MSGBodyStream:" + htmlName] != "missing" || baseline.SequenceEqual(missing))
                throw new Exception("Perda de representação HTML não foi detectada.");
            if (missing.SequenceEqual(Read(FilePath("empty.msg"))))
                throw new Exception("Stream ausente foi confundido com stream presente e vazio.");

            MustReject(FilePath("no-body.msg"));
            MustReject(FilePath("nested-only.msg"));
            MustReject(FilePath("invalid.msg"));
            Console.WriteLine("MSG: bytes/tamanho, igualdade, alteração, perda de stream, vazio e corpo fora da raiz conferidos. Sem Outlook.");
            return 0;
        }
        finally
        {
            // Exact fixture names in our own unique temporary directory only.
            foreach (string name in names) File.Delete(Path.Combine(directory, name));
            Directory.Delete(directory);
        }
    }

    private static void WriteFixture(string path, Dictionary<string, byte[]> values, byte[]? nestedBody = null)
    {
        using var root = RootStorage.Create(path);
        foreach (var entry in values)
        {
            using var stream = root.CreateStream(entry.Key);
            stream.Write(entry.Value);
        }
        if (nestedBody is not null)
        {
            // A body nested under an attachment must never substitute for the main body.
            var attachment = root.CreateStorage("__attach_version1.0_#00000000");
            using var stream = attachment.CreateStream(BodyStreams[0]);
            stream.Write(nestedBody);
        }
    }

    private static void MustReject(string path)
    {
        try { Read(path); }
        catch (InvalidDataException) { return; }
        catch (IOException) { return; }
        catch (OpenMcdf.FileFormatException) { return; }
        throw new Exception("Arquivo sem corpo MSG válido foi aceito.");
    }
}
