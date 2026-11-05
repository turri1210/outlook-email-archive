// Classification is supplementary evidence only. It never changes strict
// comparisons, the overall review status or permission to remove an original.
internal sealed class SerializedBodyEvidence
{
    public string Classification { get; set; } = "review-required";
    public bool AllKnownBodyStreamsEqual { get; set; }
    public bool CompressedRtfPresentBoth { get; set; }
    public bool CompressedRtfBytesEqual { get; set; }
    public bool HtmlBinaryPresentSource { get; set; }
    public bool HtmlBinaryPresentDestination { get; set; }
    public bool RemovalAllowed { get; } = false;
    public string Scope { get; } = "Presença e bytes nos MSGs serializados; não descomprime RTF nem comprova conteúdo no servidor, renderização ou igualdade das representações alternativas.";

    internal static SerializedBodyEvidence Compare(string sourceMsg, string destinationMsg) =>
        FromStreams(MsgBodyFingerprint.Read(sourceMsg), MsgBodyFingerprint.Read(destinationMsg));

    private static SerializedBodyEvidence FromStreams(SortedDictionary<string, string> left, SortedDictionary<string, string> right)
    {
        const string prefix = "MSGBodyStream:__substg1.0_";
        bool Present(SortedDictionary<string, string> fields, string tag) => fields.TryGetValue(prefix + tag, out var value) && value != "missing";
        bool allEqual = left.Count > 0 && left.Count == right.Count && left.All(x => right.GetValueOrDefault(x.Key) == x.Value);
        bool rtfBoth = Present(left, "10090102") && Present(right, "10090102");
        bool rtfEqual = rtfBoth && left[prefix + "10090102"] == right[prefix + "10090102"];
        return new SerializedBodyEvidence
        {
            AllKnownBodyStreamsEqual = allEqual, CompressedRtfPresentBoth = rtfBoth, CompressedRtfBytesEqual = rtfEqual,
            HtmlBinaryPresentSource = Present(left, "10130102"), HtmlBinaryPresentDestination = Present(right, "10130102"),
            Classification = allEqual ? "all-known-body-streams-equal" :
                rtfEqual ? "rtf-bytes-preserved-alternative-streams-differ" : "body-review-required"
        };
    }
    internal static void SelfTest()
    {
        const string prefix = "MSGBodyStream:__substg1.0_";
        var source = new SortedDictionary<string, string> { [prefix + "10090102"] = "length=10;sha256=AAA", [prefix + "10130102"] = "length=7;sha256=BBB" };
        var destination = new SortedDictionary<string, string>(source) { [prefix + "10130102"] = "missing" };
        var alternative = FromStreams(source, destination);
        if (!alternative.CompressedRtfBytesEqual || alternative.AllKnownBodyStreamsEqual || alternative.RemovalAllowed ||
            alternative.Classification != "rtf-bytes-preserved-alternative-streams-differ")
            throw new Exception("Evidência RTF promoveu indevidamente representações diferentes.");
        destination[prefix + "10090102"] = "length=10;sha256=CCC";
        if (FromStreams(source, destination).CompressedRtfBytesEqual) throw new Exception("RTF alterado aceito como preservado.");
        source[prefix + "10090102"] = "missing"; destination[prefix + "10090102"] = "missing";
        if (FromStreams(source, destination).CompressedRtfBytesEqual) throw new Exception("RTF ausente aceito como preservado.");
        Console.WriteLine("Classificação: RTF igual, alterado e ausente distinguidos sem aprovar remoção.");
    }
}
