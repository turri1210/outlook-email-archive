internal enum ArchiveItemState { Planned, Exported, ImportStarted, Copied, Verified, DeleteStarted, InDeletedItems, PurgeStarted, Removed, ReviewRequired, SkippedDraft, SkippedUnsupported, BackupVerifiedSourceUnmanaged, SkippedByPolicy }
internal enum ArchiveVolumeState { CreateStarted, Ready, ReviewRequired }
internal sealed class ArchiveJob
{
    public bool PlanningPending { get; set; }
    public string Id { get; set; } = "";
    public string RootDirectory { get; set; } = "";
    public string ConfigJson { get; set; } = "{}";
    public DateTimeOffset CreatedUtc { get; set; } = DateTimeOffset.UtcNow;
}
internal sealed class ArchiveItem
{
    public string Key { get; set; } = "";
    public string JobId { get; set; } = "";
    public int Version { get; set; }
    public string Token { get; set; } = Guid.NewGuid().ToString("N");
    public string? RepairToken { get; set; }
    public string? RepairManifestPath { get; set; }
    public string? RepairManifestSha256 { get; set; }
    public string? RepairPhase { get; set; }
    public bool RepairSourceUnmanaged { get; set; }
    internal string DestinationToken => RepairToken ?? Token;
    public string SourceStoreId { get; set; } = "";
    public string SourceFolderId { get; set; } = "";
    public string SourceEntryId { get; set; } = "";
    public string? SourceRecordKey { get; set; }
    public string? SourceMessageClass { get; set; }
    public string? SourceChangeKey { get; set; }
    public string? FastIdentityPath { get; set; }
    public string? FastIdentityHash { get; set; }
    public string? NativeBatchId { get; set; }
    public string? SourceRelevantHash { get; set; }
    public string? TargetRelevantHash { get; set; }
    public string[] RelativePath { get; set; } = [];
    public long EstimatedBytes { get; set; }
    public ArchiveItemState State { get; set; }
    public string? SnapshotHash { get; set; }
    public string? StageMsgPath { get; set; }
    public string? StageMsgSha256 { get; set; }
    public int? VolumeNumber { get; set; }
    public long ReservedBytes { get; set; }
    public string? TargetEntryId { get; set; }
    public string? TargetFolderId { get; set; }
    public string? TargetStoreId { get; set; }
    public string? TargetSnapshotHash { get; set; }
    public string? TargetFieldsPath { get; set; }
    public string? RemovalEntryId { get; set; }
    public string? RemovalFolderId { get; set; }
    public string? RemovalRecordKey { get; set; }
    public string? RemovalProofPath { get; set; }
    public string? RemovalProofSha256 { get; set; }
    public string? Error { get; set; }
}
internal sealed class ArchiveVolume
{
    public bool PreserveFolderStructure { get; set; }
    public Dictionary<string, ArchiveFolderDestination> FolderMappings { get; set; } = new();
    public string? ControlFolderId { get; set; }
    public bool IsExisting { get; set; }
    public string? SourceStoreId { get; set; }
    public string JobId { get; set; } = "";
    public int Number { get; set; }
    public int Version { get; set; }
    public string PstPath { get; set; } = "";
    public string? StoreId { get; set; }
    public ArchiveVolumeState State { get; set; }
    public long ObservedBytes { get; set; }
    public long InitialBytes { get; set; }
}
