using System.Globalization;
using System.Runtime.InteropServices;

internal sealed partial class OutlookArchiveBackend : IArchiveFastBatchBackend
{
    public FastMessageIdentity CaptureSourceIdentity(ArchiveItem item)
    {
        if (bulkSourceIdentities.TryGetValue(SourceBulkKey(item), out var cached)) return RequireBulkIdentity(cached);
        object? source = null;
        try
        {
            source = OpenFastSource(item, FastIdentityTags, out var values);
            return CaptureFastValues(values.Skip(FastGuardTags.Length).ToArray());
        }
        finally { Release(source); }
    }
    public FastMessageIdentity CaptureTargetIdentity(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity)
    {
        if (bulkTargetIdentities.TryGetValue(TargetBulkKey(identity), out var cached)) return RequireBulkIdentity(cached);
        object? copy = null;
        try
        {
            copy = ((dynamic)session!).GetItemFromID(identity.EntryId, volume.StoreId);
            return CaptureFast(copy);
        }
        finally { Release(copy); }
    }
    private static FastMessageIdentity CaptureFast(object message)
    {
        // One provider request; do not open recipient/attachment tables or ask
        // for bodies during provisional transfer. Full proof remains required
        // by the explicit verification/removal gate.
        return CaptureFastValues(FastProperties(message, FastIdentityTags));
    }
    private static readonly string[] FastIdentityTags = ["0x1035001F", "0x0037001F", "0x0C1F001F", "0x0C1A001F",
        "0x00390040", "0x0E060040", "0x0E04001F", "0x0E03001F", "0x0E02001F", "0x0E1B000B", "0x65E20102"];
    private static readonly string[] FastGuardTags = ["0x0FFB0102", "0x0E090102", "0x001A001F", "0x0E070003", "0x0FF90102", "0x65E20102"];
    private object OpenFastSource(ArchiveItem item, string[] extraTags, out object[] values)
    {
        object source = ((dynamic)session!).GetItemFromID(item.SourceEntryId, item.SourceStoreId);
        try
        {
            values = FastProperties(source, FastGuardTags.Concat(extraTags).ToArray());
            if (values[0] is not byte[] store || !BatchSame(Convert.ToHexString(store), item.SourceStoreId) ||
                values[1] is not byte[] folder || !BatchSame(Convert.ToHexString(folder), item.SourceFolderId) ||
                values[2] is not string messageClass || messageClass != "IPM.Note" ||
                values[3] is not int flags || flags < 0 || (flags & 0x8) != 0 ||
                item.SourceRecordKey is not null && (values[4] is not byte[] key || Convert.ToHexString(key) != item.SourceRecordKey))
                throw new ArchiveReviewException("Origem mudou, é rascunho ou está fora do escopo.");
            return source;
        }
        catch { Release(source); throw; }
    }
    private static FastMessageIdentity CaptureFastValues(object[] values)
    {
        var fields = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Version"] = "fast-metadata-v2",
            ["InternetMessageId"] = FastText(values[0]), ["Subject"] = FastText(values[1]),
            ["Sender"] = FastText(values[2]), ["SenderName"] = FastText(values[3]),
            ["Sent"] = FastDate(values[4]), ["Received"] = FastDate(values[5]),
            ["To"] = FastText(values[6]), ["Cc"] = FastText(values[7]), ["Bcc"] = FastText(values[8]),
            ["HasAttachments"] = values[9] is bool hasAttachments ? hasAttachments.ToString() : throw new ArchiveReviewException("Indicador de anexos indisponível para identificação rápida.")
        };
        if (fields["Sent"].Length == 0 && fields["Received"].Length == 0)
            throw new ArchiveReviewException("Datas indisponíveis para identificação rápida.");
        return new(fields, values[10] is byte[] changeKey ? Convert.ToHexString(changeKey) : null);
    }
    private static object[] FastProperties(object message, string[] tags)
    {
        object? accessor = null;
        try
        {
            accessor = ((dynamic)message).PropertyAccessor;
            object result = ((dynamic)accessor).GetProperties(tags.Select(x => "http://schemas.microsoft.com/mapi/proptag/" + x).ToArray());
            if (result is not Array array || array.Length != tags.Length)
                throw new ArchiveReviewException("Resposta incompleta de propriedades do Outlook.");
            return array.Cast<object>().ToArray();
        }
        finally { Release(accessor); }
    }
    private static bool FastMissing(object? value) => value is null || value is ErrorWrapper wrapper && wrapper.ErrorCode == unchecked((int)0x8004010F) || value is int error && error == unchecked((int)0x8004010F);
    private static string FastText(object? value) => value is string text ? text : FastMissing(value) ? "" : throw new ArchiveReviewException("Texto de identificação inválido.");
    private static string FastDate(object? value) => value is DateTime date
        ? DateTime.SpecifyKind(date, DateTimeKind.Utc).ToString("O", CultureInfo.InvariantCulture)
        : FastMissing(value) ? "" : throw new ArchiveReviewException("Data de identificação inválida.");
}
