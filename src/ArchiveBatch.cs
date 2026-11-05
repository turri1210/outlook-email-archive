internal enum ArchiveBatchState { Prepared, CopyIntent, Matched, Complete, ReviewRequired }
internal sealed record BatchFolder(string StoreId, string FolderId);
internal interface IArchiveBatchBackend
{
    bool CanBatch(ArchiveItem item);
    BatchFolder PrepareBatch(ArchiveVolume volume, ArchiveBatch batch);
    void CopyBatch(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> items);
    void AssertBatchIdle(ArchiveBatch batch);
    List<DestinationIdentity> BatchCandidates(ArchiveVolume volume, ArchiveBatch batch);
    void MarkBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem item, DestinationIdentity identity);
    DestinationIdentity FinalizeBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem item, DestinationIdentity identity);
    void CloseBatch(ArchiveVolume volume, ArchiveBatch batch);
}
internal sealed class ArchiveBatch
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string JobId { get; set; } = "";
    public string DirectoryPath { get; set; } = "";
    public int Version { get; set; }
    public int VolumeNumber { get; set; }
    public string SourceStoreId { get; set; } = "";
    public string SourceFolderId { get; set; } = "";
    public string[] RelativePath { get; set; } = [];
    public string? StoreId { get; set; }
    public string? FolderId { get; set; }
    public List<string> Items { get; set; } = [];
    public bool FastCopy { get; set; }
    public string ManifestSha256 { get; set; } = "";
    public ArchiveBatchState State { get; set; }
    public Dictionary<string, DestinationIdentity> Matches { get; set; } = new(StringComparer.Ordinal);
    public string? Error { get; set; }
}
