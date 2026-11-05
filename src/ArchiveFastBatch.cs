internal sealed record FastMessageIdentity(Dictionary<string, string> Fields, string? ChangeKey = null)
{
    public string Hash => ArchiveFiles.HashFields(Fields);
}
internal interface IArchiveFastBatchBackend
{
    FastMessageIdentity CaptureSourceIdentity(ArchiveItem item);
    FastMessageIdentity CaptureTargetIdentity(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity);
}
internal interface IArchiveBulkFastBatchBackend : IArchiveFastBatchBackend
{
    void RefreshSourceIdentities(IReadOnlyList<ArchiveItem> rows);
    void RefreshTargetIdentities(ArchiveVolume volume, IReadOnlyList<DestinationIdentity> candidates);
    void ClearFastIdentityCache();
}
