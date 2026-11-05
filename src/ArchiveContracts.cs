using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

internal sealed record SourceDescriptor(string StoreId, string FolderId, string EntryId, string[] RelativePath,
    long Size, bool IsSupported, bool IsSent, string? RecordKey, DateTime? MessageDateUtc = null, string? DateBasis = null, string? MessageClass = null);
internal sealed record FolderScopeSnapshot(string FolderId, string[] RelativePath, List<string> EntryIds);
internal sealed record ArrivalCheckResult(int NewItemsCount, DateTimeOffset CheckedUtc, string? Error = null);
internal sealed record DestinationIdentity(string StoreId, string FolderId, string EntryId);
internal sealed record TokenLookup(List<DestinationIdentity> Matches, int UnmarkedItems);
internal sealed record DiskRoundtrip(bool Reopened, bool ExclusiveDiskRead);
internal sealed record SourceLocation(string EntryId, string FolderId, bool InDeletedItems, string? RecordKey);
internal sealed record ArchiveSnapshot(Dictionary<string, string> Fields, string MessagePath, string MessageSha256,
    string? ChangeKey = null, Dictionary<string, string>? RelevantFields = null)
{
    public string Hash => ArchiveFiles.HashFields(Fields);
}
internal interface IArchiveBackend : IDisposable
{
    List<SourceDescriptor> Discover(ArchiveConfig config);
    List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config);
    List<SourceDescriptor> DiscoverFrozen(ArchiveConfig config, IReadOnlyList<FolderScopeSnapshot> scope);
    ArchiveSnapshot Export(ArchiveItem item, string evidenceDirectory);
    ArchiveSnapshot SourceSnapshot(ArchiveItem item, SourceLocation? location, string evidenceDirectory);
    bool SourceMatchesChangeKey(ArchiveItem item, string expectedChangeKey);
    ArchiveSnapshot BackupSnapshot(ArchiveItem item, string evidenceDirectory);
    void CreateVolume(ArchiveVolume volume);
    void AttachVolume(ArchiveVolume volume);
    long MeasureVolume(ArchiveVolume volume);
    TokenLookup FindToken(ArchiveVolume volume, ArchiveItem item);
    DestinationIdentity Import(ArchiveVolume volume, ArchiveItem item);
    ArchiveSnapshot TargetSnapshot(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity, string evidenceDirectory, bool requireToken = true);
    List<DestinationIdentity> UnmarkedCandidates(ArchiveVolume volume, ArchiveItem item);
    void MarkCopy(ArchiveVolume volume, ArchiveItem item, DestinationIdentity identity);
    DiskRoundtrip ReopenVolume(ArchiveVolume volume);
    SourceLocation? LocateSource(ArchiveItem item);
    SourceLocation MoveToDeletedItems(ArchiveItem item, string? expectedChangeKey);
    void DeleteFromDeletedItems(ArchiveItem item, SourceLocation location, string? expectedChangeKey);
}
internal interface IArchiveResumableScopeBackend
{
    List<FolderScopeSnapshot> CaptureScope(ArchiveConfig config,IReadOnlyList<FolderScopeSnapshot> saved,Action<FolderScopeSnapshot> checkpoint);
}
internal sealed class ArchiveReviewException(string message) : Exception(message);
internal sealed class ArchiveInjectedCrashException(string message) : Exception(message);

internal sealed class ArchiveConfig
{
    public bool PreserveFolderStructure { get; set; }
    public Dictionary<string, ArchiveFolderDestination> FolderMappings { get; set; } = new();
    public string RootDirectory { get; set; } = "";
    public string SourceStoreId { get; set; } = "";
    public string SourceFolderId { get; set; } = "";
    public string? SourceLabel { get; set; }
    public string? ExistingPstPath { get; set; }
    public string? DestinationPstPath { get; set; }
    public bool IncludeSubfolders { get; set; } = true;
    public DateOnly? CutoffDateInclusive { get; set; }
    public string CutoffTimeZoneId { get; set; } = TimeZoneInfo.Local.Id;
    public int MaxItems { get; set; } = 1_000_000;
    public long TargetVolumeBytes { get; set; } = ArchiveVolumePolicy.DefaultTargetVolumeBytes;
    public long HardLimitBytes { get; set; } = ArchiveVolumePolicy.DefaultHardLimitBytes;
    public long ReserveBytes { get; set; } = ArchiveVolumePolicy.DefaultReserveBytes;
    public string? SyntheticSourcePstPath { get; set; }
    public bool AllowRetainedMsgBackupForRemoval { get; set; }
    public bool MoveValidatedToDeletedItems { get; set; }
    public void Validate()
    {
        ArchiveFiles.ValidateLocalDirectory(RootDirectory);
        if (CutoffDateInclusive is { } cutoff)
        {
            if (cutoff.Year < 1601) throw new ArgumentException("A data de corte deve ser a partir de 01/01/1601.");
            _ = TimeZoneInfo.FindSystemTimeZoneById(CutoffTimeZoneId);
        }
        if (ExistingPstPath is not null) ArchiveFiles.ValidateExistingPst(ExistingPstPath);
        if(DestinationPstPath is not null)
        {
            if(!Path.IsPathFullyQualified(DestinationPstPath) || !string.Equals(Path.GetExtension(DestinationPstPath),".pst",StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Informe um caminho completo com extensão .pst para o destino.");
            ArchiveFiles.ValidateLocalDirectory(Path.GetDirectoryName(Path.GetFullPath(DestinationPstPath))!);
            if(ExistingPstPath is not null && !string.Equals(Path.GetFullPath(ExistingPstPath),Path.GetFullPath(DestinationPstPath),StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("O PST existente não corresponde ao destino escolhido.");
        }
        if (Path.GetFullPath(RootDirectory).TrimEnd(Path.DirectorySeparatorChar).Length > 120)
            throw new ArgumentException("Escolha uma raiz com até 120 caracteres para compatibilidade dos caminhos MSG com o Outlook clássico.");
        if (string.IsNullOrWhiteSpace(SourceStoreId) || string.IsNullOrWhiteSpace(SourceFolderId) || MaxItems is < 1 or > 1_000_000)
            throw new ArgumentException("Origem/limite inválidos.");
        var policy = ArchiveVolumePolicy.Decide(ArchiveVolumePolicy.DefaultEmptyVolumeBaseBytes, 0, 0, TargetVolumeBytes, HardLimitBytes, ReserveBytes);
        if (!policy.Allowed) throw new ArgumentException("Política de volumes inválida: " + policy.Reason);
    }
}
internal sealed record ArchiveFolderDestination(int? DefaultFolderType, string DefaultFolderName, string[] Subfolders);
internal sealed record ArchivePlanEnvelope(ArchiveConfig Config, string PlanSha256, string? InitialScopeSha256 = null);
internal static class ArchiveFiles
{
    internal static void ValidateExistingPst(string path)
    {
        if (!Path.IsPathFullyQualified(path) || !string.Equals(Path.GetExtension(path), ".pst", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Selecione um arquivo PST existente com caminho completo.");
        ValidateLocalDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        if (!File.Exists(path)) throw new FileNotFoundException("O PST selecionado não existe; ele não será criado automaticamente.", path);
        if (File.GetAttributes(path).HasFlag(FileAttributes.ReparsePoint)) throw new ArgumentException("O PST não pode ser um redirecionamento.");
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        if (input.Length < 12) throw new ArgumentException("O PST selecionado está incompleto.");
        // Outlook locks byte ranges of an attached PST. A brief read-only map
        // reads its immutable format header without requesting those ranges.
        using var mapping = System.IO.MemoryMappedFiles.MemoryMappedFile.CreateFromFile(input, null, 0, System.IO.MemoryMappedFiles.MemoryMappedFileAccess.Read, HandleInheritability.None, true);
        using var view = mapping.CreateViewAccessor(0, 12, System.IO.MemoryMappedFiles.MemoryMappedFileAccess.Read);
        byte[] header = new byte[12]; view.ReadArray(0, header, 0, header.Length);
        if (!header.AsSpan(0, 4).SequenceEqual(new byte[] { 0x21, 0x42, 0x44, 0x4E }) || header[8] != 0x53 || header[9] != 0x4D)
            throw new ArgumentException("O arquivo selecionado não tem cabeçalho PST reconhecido.");
        ushort version = System.Buffers.Binary.BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(10));
        if (version is not (23 or 36)) throw new ArgumentException("Use um PST Unicode compatível (versão 23 ou 36). PST ANSI/protegido não é suportado nesta opção.");
    }
    internal static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    internal static string Sha(string path)
    { if(ArchiveEvidenceLease.TryHash(path,out string cached))return cached;using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite); return Convert.ToHexString(SHA256.HashData(stream)); }
    internal static string HashFields(IReadOnlyDictionary<string, string> fields) => Convert.ToHexString(SHA256.HashData(
        Encoding.UTF8.GetBytes(JsonSerializer.Serialize(new SortedDictionary<string, string>(fields.ToDictionary(x => x.Key, x => x.Value), StringComparer.Ordinal)))));
    internal static Dictionary<string, string> ReadFields(string path) => JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(path))
        ?? throw new InvalidDataException("Snapshot inválido.");
    internal static void WriteNew<T>(string path, T data)
    {
        string temp = path + "." + Guid.NewGuid().ToString("N") + ".partial";
        using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
        { JsonSerializer.Serialize(stream, data, JsonOptions); stream.Flush(true); }
        File.Move(temp, path, false);
    }
    internal static void ValidateLocalDirectory(string path)
    {
        if (!Path.IsPathFullyQualified(path) || path.StartsWith(@"\\", StringComparison.Ordinal)) throw new ArgumentException("Destino absoluto local é obrigatório.");
        string full = Path.GetFullPath(path);
        if (full.TrimEnd(Path.DirectorySeparatorChar) == Path.GetPathRoot(full)!.TrimEnd(Path.DirectorySeparatorChar))
            throw new ArgumentException("Escolha uma pasta própria, não a raiz do disco.");
        if (new DriveInfo(Path.GetPathRoot(full)!).DriveType != DriveType.Fixed) throw new ArgumentException("Use disco local fixo.");
        for (var d = new DirectoryInfo(full); d is not null; d = d.Parent)
            if (Directory.Exists(d.FullName) && (d.Attributes.HasFlag(FileAttributes.ReparsePoint) || d.Name.Contains("OneDrive", StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException("Destino não pode passar por OneDrive ou redirecionamento.");
        foreach (string name in new[] { "OneDrive", "OneDriveCommercial", "OneDriveConsumer" })
        {
            string? root = Environment.GetEnvironmentVariable(name);
            if (!string.IsNullOrEmpty(root) && PilotCommand.IsWithin(full, root)) throw new ArgumentException("Destino não pode ficar no OneDrive.");
        }
    }
    internal static string Child(string root, string path)
    {
        string full = Path.GetFullPath(path);
        if (!PilotCommand.IsWithin(full, root) || string.Equals(full, Path.GetFullPath(root), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Caminho fora da execução.");
        ValidateLocalDirectory(Path.GetDirectoryName(full)!); return full;
    }
}
