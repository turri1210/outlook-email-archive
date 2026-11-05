using System.Buffers.Binary;
using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OpenMcdf;

// Offline diagnostic only. Never renders HTML, fetches a URL, opens Outlook,
// changes a fingerprint policy, or declares a difference safe to ignore.
internal static class MsgDifferenceAnalysis
{
    internal static void SelfTest()
    {
        var lineEndings = JsonSerializer.SerializeToElement(CompareText("Olá\r\n", "Olá\n"));
        if (lineEndings.GetProperty("ExactEqual").GetBoolean() ||
            !lineEndings.GetProperty("EqualAfterLineEndingNormalization").GetBoolean() ||
            !lineEndings.GetProperty("DiagnosticOnly").GetBoolean())
            throw new Exception("Diagnóstico de quebras de linha inválido.");
        var changed = JsonSerializer.SerializeToElement(CompareText("Total: 100", "Total: 900"));
        if (changed.GetProperty("NonWhitespaceCodePointsEqual").GetBoolean())
            throw new Exception("Alteração de conteúdo ignorada.");
        var links = JsonSerializer.SerializeToElement(CompareHtml("<a href='https://example.invalid/a'>Texto</a>", "<a href='https://example.invalid/b'>Texto</a>"));
        if (links.GetProperty("HyperlinkTargets").GetProperty("Equal").GetBoolean())
            throw new Exception("Mudança no destino de link ignorada.");
        var missing = JsonSerializer.SerializeToElement(CompareHtml("<p>Olá</p>", null));
        if (missing.GetProperty("Right").ValueKind != JsonValueKind.Null ||
            missing.GetProperty("HyperlinkTargets").ValueKind != JsonValueKind.Null)
            throw new Exception("HTML ausente tratado como comparação aprovada.");
        Console.WriteLine("Análise: quebras de linha, conteúdo alterado, URLs diferentes e HTML ausente distinguidos sem aprovação automática.");
    }

    private const int MaxStreamBytes = 32 * 1024 * 1024;
    private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(3);
    private static readonly string[] Tags = ["1000001F", "1000001E", "10130102", "1013001F", "10090102"];

    internal static int Execute(string[] args)
    {
        if (args.Length != 2)
        { Console.Error.WriteLine("Uso: msg-diff <origem.msg> <destino.msg>"); return 2; }
        try
        {
            Console.WriteLine(JsonSerializer.Serialize(Analyze(args[0], args[1]), new JsonSerializerOptions { WriteIndented = true }));
            return 0;
        }
        catch (Exception ex)
        {
            // Exception messages may contain message-derived data; emit a type only.
            Console.Error.WriteLine("Análise MSG indisponível: " + ex.GetType().Name);
            return 1;
        }
    }

    internal static object Analyze(string left, string right)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var l = RootStorage.OpenRead(left);
        using var r = RootStorage.OpenRead(right);
        var leftBytes = Tags.ToDictionary(tag => tag, tag => ReadStream(l, "__substg1.0_" + tag));
        var rightBytes = Tags.ToDictionary(tag => tag, tag => ReadStream(r, "__substg1.0_" + tag));
        var streams = Tags.Select(tag => new
        {
            Tag = tag,
            Left = ByteStats(leftBytes[tag]), Right = ByteStats(rightBytes[tag]),
            Equal = SameBytes(leftBytes[tag], rightBytes[tag])
        }).ToArray();
        var lb = DecodeBody(leftBytes, l);
        var rb = DecodeBody(rightBytes, r);
        var lh = DecodeHtml(leftBytes, l);
        var rh = DecodeHtml(rightBytes, r);
        return new
        {
            SchemaVersion = 1,
            Scope = "Comparação offline dos streams serializados em dois MSGs; não verifica o servidor, renderização ou todas as propriedades MAPI.",
            Streams = streams,
            Body = CompareDecoded(lb, rb),
            Html = new
            {
                Decoded = CompareDecoded(lh, rh),
                Structure = CompareHtml(lh?.Text, rh?.Text),
                Limits = "Extração estática heurística de tags e atributos. Não valida DOM, CSS, scripts, recursos remotos, posicionamento ou aparência. Igualdade de URLs não comprova igualdade das imagens."
            },
            Rtf = new { ComparedAsCompressedBytesOnly = true, Decompressed = false },
            Policy = "Diferenças permanecem registradas; nenhuma normalização diagnóstica altera a aprovação do piloto."
        };
    }

    internal static string ReadHtmlForRestore(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var root=RootStorage.OpenRead(path);
        var streams=new Dictionary<string,byte[]?>{["10130102"]=ReadStream(root,"__substg1.0_10130102"),["1013001F"]=ReadStream(root,"__substg1.0_1013001F")};
        var decoded=DecodeHtml(streams,root) ?? throw new InvalidDataException("HTML original ausente.");
        if(decoded.Assumed)throw new InvalidDataException("Codificação do HTML não foi declarada; restauração automática bloqueada.");
        return decoded.Text;
    }

    private static byte[]? ReadStream(RootStorage root, string name)
    {
        if (!root.TryOpenStream(name, out var stream)) return null;
        using (stream)
        {
            if (stream.Length > MaxStreamBytes) throw new InvalidDataException("Stream excede o limite de análise.");
            var bytes = new byte[checked((int)stream.Length)];
            stream.ReadExactly(bytes);
            return bytes;
        }
    }

    private static object ByteStats(byte[]? bytes) => bytes is null
        ? new { Present = false, Length = (int?)null, Sha256 = (string?)null }
        : new { Present = true, Length = (int?)bytes.Length, Sha256 = (string?)Hash(bytes) };
    private static bool SameBytes(byte[]? left, byte[]? right) => left is null ? right is null : right is not null && left.AsSpan().SequenceEqual(right);
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static string TextHash(string text) => Hash(Encoding.UTF8.GetBytes(text));

    private sealed record Decoded(string Text, string Tag, string Encoding, string Selection, bool Assumed);
    private static Decoded? DecodeBody(Dictionary<string, byte[]?> streams, RootStorage root)
    {
        if (streams["1000001F"] is byte[] unicode)
            return new(new UnicodeEncoding(false, false, true).GetString(unicode), "1000001F", "utf-16LE", "property-type", false);
        if (streams["1000001E"] is byte[] ansi)
        {
            int? cp = CodePage(root, 0x3FFD0003) ?? CodePage(root, 0x3FDE0003);
            var encoding = Encoding.GetEncoding(cp ?? 1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            return new(encoding.GetString(ansi), "1000001E", encoding.WebName, cp.HasValue ? "message-codepage" : "windows-1252-fallback", !cp.HasValue);
        }
        return null;
    }

    private static Decoded? DecodeHtml(Dictionary<string, byte[]?> streams, RootStorage root)
    {
        if (streams["10130102"] is byte[] bytes)
        {
            Encoding encoding;
            string selection;
            bool assumed = false;
            int skip = 0;
            if (bytes.AsSpan().StartsWith(new byte[] { 0xEF, 0xBB, 0xBF }))
            { encoding = new UTF8Encoding(false, true); selection = "utf8-bom"; skip = 3; }
            else if (bytes.AsSpan().StartsWith(new byte[] { 0xFF, 0xFE }))
            { encoding = new UnicodeEncoding(false, false, true); selection = "utf16le-bom"; skip = 2; }
            else if (bytes.AsSpan().StartsWith(new byte[] { 0xFE, 0xFF }))
            { encoding = new UnicodeEncoding(true, false, true); selection = "utf16be-bom"; skip = 2; }
            else
            {
                string header = Encoding.Latin1.GetString(bytes, 0, Math.Min(bytes.Length, 16384));
                string? declared = null;
                foreach (Match meta in Matches(header, @"<meta\b(?:""[^""]*""|'[^']*'|[^'"">])*>"))
                {
                    Match charset = Match(meta.Value, @"\bcharset\s*=\s*[""']?\s*(?<value>[A-Za-z0-9_.-]+)");
                    if (charset.Success) { declared = charset.Groups["value"].Value; break; }
                }
                if (declared is not null)
                { encoding = Encoding.GetEncoding(declared, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); selection = "html-meta-charset"; }
                else if ((CodePage(root, 0x3FDE0003) ?? CodePage(root, 0x3FFD0003)) is int cp)
                { encoding = Encoding.GetEncoding(cp, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); selection = "mapi-codepage"; }
                else
                {
                    assumed = true;
                    try { new UTF8Encoding(false, true).GetString(bytes); encoding = new UTF8Encoding(false, true); selection = "valid-utf8-assumption"; }
                    catch (DecoderFallbackException)
                    { encoding = Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback); selection = "windows-1252-fallback"; }
                }
            }
            return new(encoding.GetString(bytes, skip, bytes.Length - skip), "10130102", encoding.WebName, selection, assumed);
        }
        if (streams["1013001F"] is byte[] unicode)
            return new(new UnicodeEncoding(false, false, true).GetString(unicode), "1013001F", "utf-16LE", "property-type", false);
        return null;
    }

    private static int? CodePage(RootStorage root, uint propertyTag)
    {
        var bytes = ReadStream(root, "__properties_version1.0");
        if (bytes is null || bytes.Length < 32) return null;
        // Top-level MSG property stream: 32-byte header followed by 16-byte entries.
        for (int offset = 32; offset <= bytes.Length - 16; offset += 16)
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4)) == propertyTag)
                return BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(offset + 8, 4));
        return null;
    }

    private static object CompareDecoded(Decoded? left, Decoded? right) => new
    {
        LeftPresent = left is not null, RightPresent = right is not null,
        LeftDecode = left is null ? null : new { left.Tag, left.Encoding, left.Selection, left.Assumed },
        RightDecode = right is null ? null : new { right.Tag, right.Encoding, right.Selection, right.Assumed },
        Comparison = left is null || right is null ? null : CompareText(left.Text, right.Text)
    };

    private static object CompareText(string left, string right)
    {
        Rune[] l = left.EnumerateRunes().ToArray(), r = right.EnumerateRunes().ToArray();
        int prefix = 0, suffix = 0;
        while (prefix < Math.Min(l.Length, r.Length) && l[prefix] == r[prefix]) prefix++;
        while (suffix < Math.Min(l.Length - prefix, r.Length - prefix) && l[l.Length - 1 - suffix] == r[r.Length - 1 - suffix]) suffix++;
        var lSpan = l.Skip(prefix).Take(l.Length - prefix - suffix).ToArray();
        var rSpan = r.Skip(prefix).Take(r.Length - prefix - suffix).ToArray();
        return new
        {
            ExactEqual = string.Equals(left, right, StringComparison.Ordinal),
            Left = TextStats(left, l), Right = TextStats(right, r),
            CommonPrefixCodePoints = prefix, CommonSuffixCodePoints = suffix,
            DifferingEnvelope = new
            {
                LeftCodePoints = lSpan.Length, RightCodePoints = rSpan.Length,
                LeftFirstEight = lSpan.Take(8).Select(CodePoint).ToArray(),
                RightFirstEight = rSpan.Take(8).Select(CodePoint).ToArray(),
                LeftEnvelopeWhitespaceOnly = lSpan.All(Rune.IsWhiteSpace),
                RightEnvelopeWhitespaceOnly = rSpan.All(Rune.IsWhiteSpace),
                Note = "Envelope delimitado por prefixo/sufixo comuns; pode conter trechos internos iguais. Não é alinhamento mínimo."
            },
            NonWhitespaceCodePointsEqual = l.Where(x => !Rune.IsWhiteSpace(x)).SequenceEqual(r.Where(x => !Rune.IsWhiteSpace(x))),
            EqualAfterWhitespaceCollapse = NormalizeWhitespace(l) == NormalizeWhitespace(r),
            EqualAfterLineEndingNormalization = NormalizeLines(left) == NormalizeLines(right),
            DiagnosticOnly = true
        };
    }

    private static object TextStats(string text, Rune[] runes) => new
    {
        Utf16CodeUnits = text.Length, CodePoints = runes.Length, Utf8Sha256 = TextHash(text),
        WhitespaceCodePoints = runes.Count(Rune.IsWhiteSpace),
        WhitespaceAndNullCounts = runes.Where(x => Rune.IsWhiteSpace(x) || x.Value == 0)
            .GroupBy(x => x.Value).OrderBy(x => x.Key)
            .Select(g => new { CodePoint = "U+" + g.Key.ToString("X4", CultureInfo.InvariantCulture), Count = g.Count() }).ToArray()
    };
    private static object CodePoint(Rune rune) => new
    { Unicode = "U+" + rune.Value.ToString("X4", CultureInfo.InvariantCulture), Category = Rune.GetUnicodeCategory(rune).ToString() };
    private static string NormalizeLines(string text) => text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
    private static string NormalizeWhitespace(IEnumerable<Rune> runes)
    {
        var result = new StringBuilder();
        bool whitespace = false;
        foreach (Rune rune in runes)
        {
            if (Rune.IsWhiteSpace(rune))
            { if (!whitespace) result.Append(' '); whitespace = true; }
            else { result.Append(rune.ToString()); whitespace = false; }
        }
        return result.ToString();
    }

    private static object CompareHtml(string? left, string? right)
    {
        var l = left is null ? null : HtmlStats(left);
        var r = right is null ? null : HtmlStats(right);
        return new
        {
            Left = l?.Summary, Right = r?.Summary,
            HyperlinkTargets = l is null || r is null ? null : CompareMultiset(l.Links, r.Links),
            ImageSources = l is null || r is null ? null : CompareMultiset(l.Images, r.Images),
            ExtractedText = left is null || right is null ? null : CompareText(ExtractText(left), ExtractText(right)),
            ExtractedTextLimits = "Texto heurístico: remove comentários/script/style e tags, decodifica entidades; não calcula visibilidade, CSS ou espaços do navegador."
        };
    }

    private sealed record HtmlData(object Summary, List<string> Links, List<string> Images);
    private static HtmlData HtmlStats(string html)
    {
        var links = new List<string>(); var images = new List<string>();
        var tags = new SortedDictionary<string, int>(StringComparer.Ordinal);
        var noComments = Replace(html, @"<!--[\s\S]*?-->", "");
        foreach (Match tag in Matches(noComments, @"<\s*(?<name>[A-Za-z][A-Za-z0-9:_-]*)\b(?<attrs>(?:""[^""]*""|'[^']*'|[^'"">])*)>"))
        {
            string name = tag.Groups["name"].Value.ToLowerInvariant();
            tags[name] = tags.GetValueOrDefault(name) + 1;
            var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match attr in Matches(tag.Groups["attrs"].Value, @"(?<name>[A-Za-z_:][A-Za-z0-9:_.-]*)\s*=\s*(?:""(?<double>[^""]*)""|'(?<single>[^']*)'|(?<bare>[^\s""'=<>`]+))"))
            {
                string value = attr.Groups["double"].Success ? attr.Groups["double"].Value :
                    attr.Groups["single"].Success ? attr.Groups["single"].Value : attr.Groups["bare"].Value;
                attributes.TryAdd(attr.Groups["name"].Value, WebUtility.HtmlDecode(value));
            }
            if (name is "a" or "area" && attributes.TryGetValue("href", out string? href)) links.Add(href);
            if (name == "img")
                foreach (string attribute in new[] { "src", "srcset" })
                    if (attributes.TryGetValue(attribute, out string? src)) images.Add(attribute + ":" + src);
        }
        return new(new
        {
            OpeningTagCount = tags.Values.Sum(), TagCounts = tags,
            MicrosoftWordNamespace = html.Contains("urn:schemas-microsoft-com:office:word", StringComparison.OrdinalIgnoreCase),
            MicrosoftOfficeNamespace = html.Contains("urn:schemas-microsoft-com:office:office", StringComparison.OrdinalIgnoreCase),
            MicrosoftWordMarker = Match(html, @"Microsoft\s+(?:Office\s+)?Word").Success,
            MsoTokenCount = Matches(html, @"\bmso[-:]").Count,
            HyperlinkTargetCount = links.Count, ImageSourceCount = images.Count
        }, links, images);
    }

    private static object CompareMultiset(List<string> left, List<string> right)
    {
        var l = left.Select(TextHash).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var r = right.Select(TextHash).GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var differences = l.Keys.Concat(r.Keys).Distinct().Order(StringComparer.Ordinal)
            .Where(key => l.GetValueOrDefault(key) != r.GetValueOrDefault(key))
            .Select(key => new { ValueSha256 = key, LeftCount = l.GetValueOrDefault(key), RightCount = r.GetValueOrDefault(key) }).ToArray();
        return new { LeftCount = left.Count, RightCount = right.Count, Equal = differences.Length == 0, Differences = differences };
    }

    private static string ExtractText(string html)
    {
        html = Replace(html, @"<!--[\s\S]*?-->", "");
        html = Replace(html, @"<(script|style)\b[^>]*>[\s\S]*?</\1\s*>", "");
        html = Replace(html, @"<(?:""[^""]*""|'[^']*'|[^'"">])*>", "");
        return WebUtility.HtmlDecode(html);
    }
    private static Match Match(string input, string pattern) => Regex.Match(input, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, RegexTimeout);
    private static MatchCollection Matches(string input, string pattern) => Regex.Matches(input, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, RegexTimeout);
    private static string Replace(string input, string pattern, string replacement) => Regex.Replace(input, pattern, replacement, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, RegexTimeout);
}
