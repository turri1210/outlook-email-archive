using System.Text.Json;

internal sealed partial class ArchiveEngine
{
    // Explicit recovery only. A stable intent survives crashes; old rows and
    // snapshots remain in the immutable manifest. Never delete an old copy.
    internal void ResolvePending(string jobId, IReadOnlySet<string> keys, bool allowRetainedBackup)
    {
        var job = ledger.GetJob(jobId); var config = Envelope(job).Config;
        var reports = new List<object>();
        foreach (var row in ledger.GetItems(jobId).Where(x => keys.Contains(x.Key) && x.State == ArchiveItemState.ReviewRequired))
        {
            Console.WriteLine($"Revisão assistida: {row.Key[..12]} — comparando origem, backup e PST…");
            var old = JsonSerializer.Deserialize<ArchiveItem>(JsonSerializer.Serialize(row))!;
            string? sourcePath = null, targetPath = null; string decision = "revisão";
            try
            {
                var location = backend.LocateSource(row);
                bool unknown = location is null || location.FolderId != row.SourceFolderId;
                ArchiveSnapshot? source = null;
                if (!unknown) { source = backend.Export(row, Evidence(job,row,"comparison-source")); sourcePath = source.MessagePath; }
                CheckBackupFile(job,row);
                var backup = backend.BackupSnapshot(row,Evidence(job,row,"comparison-backup"));
                if (!RecordedBackupMatches(job,row,backup)) throw new ArchiveReviewException("Backup não corresponde à evidência preservada.");
                ArchiveVolume? volume = row.VolumeNumber is null ? null : Volume(row);
                if(volume is null && (row.TargetEntryId is not null || row.RepairToken is not null))
                    throw new ArchiveReviewException("Destino registrado sem volume; revisão necessária.");
                ArchiveSnapshot? target = null; DestinationIdentity? identity = null;
                if(volume is not null)
                {
                EnsureVolume(volume);
                var found = backend.FindToken(volume,row);
                if (found.Matches.Count > 1) throw new ArchiveReviewException("Mais de uma cópia com a mesma identificação; nenhuma será substituída.");
                identity = found.Matches.SingleOrDefault();
                if (identity is null && row.TargetEntryId is not null)
                    identity = new(volume.StoreId!, row.TargetFolderId!, row.TargetEntryId);
                if (identity is not null)
                {
                    try
                    {
                        target = backend.TargetSnapshot(volume,row,identity,Evidence(job,row,"comparison-target"),false);
                        targetPath = target.MessagePath;
                    }
                    catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x8004010F))
                    { identity = null; }
                }
                }
                var desired = source ?? backup;
                bool oldCopyEquivalent = target is not null && RelevantMessageValidation.Equivalent(desired,target);
                bool sourceMatchesBackup = source is null || RelevantMessageValidation.Equivalent(backup,source);
                ArchiveFiles.WriteNew(Path.Combine(RepairDirectory(job,row,"comparison-table"),"comparison.json"), new
                {
                    Key=row.Key, SourceLocated=!unknown, SourceMsg=sourcePath, OriginalBackup=row.StageMsgPath, PreviousTargetMsg=targetPath,
                    SourceMatchesBackup=sourceMatchesBackup, PreviousTargetMatches=oldCopyEquivalent,
                    BackupProfile=backup.RelevantFields, SourceProfile=source?.RelevantFields, PreviousTargetProfile=target?.RelevantFields,
                    Differences=target is null ? "Cópia não localizada" : RelevantMessageValidation.DifferenceSummary(desired,target)
                });
                if (unknown && oldCopyEquivalent && row.RepairToken is null)
                {
                    backend.MarkCopy(volume!,row,identity!);
                    ConfirmBackupOnly(job,row,volume!,identity!,allowRetainedBackup,location);
                    decision="PST validado; origem não localizada";
                }
                else
                {
                    bool adopt = oldCopyEquivalent && sourceMatchesBackup && row.RepairToken is null;
                    RepairCopy(job,row,source,unknown,adopt ? identity : null,config);
                    decision = adopt ? "Cópia existente aprovada" : "Nova cópia identificada; anterior preservada";
                    if (unknown) ConfirmBackupOnly(job,row,Volume(row),Unique(Volume(row),row),allowRetainedBackup,location);
                }
            }
            catch (ArchiveInjectedCrashException) { throw; }
            catch (Exception ex) { Review(row,"repair-pending:"+ex.Message); decision="Revisão: "+ex.Message; }
            reports.Add(new { Key=row.Key, PreviousState=old.State.ToString(), State=row.State.ToString(), Decision=decision,
                SourceMsg=sourcePath, PreviousTargetMsg=targetPath, RepairManifest=row.RepairManifestPath });
        }
        string report = Path.Combine(job.RootDirectory,"recovery-result-"+Guid.NewGuid().ToString("N")[..12]+".json");
        ArchiveFiles.WriteNew(report,reports); Console.WriteLine("RELATORIO="+report); Status(jobId);
    }

    private void RepairCopy(ArchiveJob job, ArchiveItem row, ArchiveSnapshot? source, bool sourceUnknown, DestinationIdentity? adopt, ArchiveConfig config)
    {
        bool freshIntent = false;
        if (row.RepairToken is null)
        {
            var previous = JsonSerializer.Deserialize<ArchiveItem>(JsonSerializer.Serialize(row))!;
            string folder = RepairDirectory(job,row,"repair-source");
            if (source is null)
            {
                CheckBackupFile(job,row);
                var replay = backend.BackupSnapshot(row,folder);
                if (!RecordedBackupMatches(job,row,replay)) throw new ArchiveReviewException("Backup íntegro não confirmado para recopiar.");
                string copied = Path.Combine(folder,"message.msg"); File.Copy(row.StageMsgPath!,copied,false);
                source = replay with { MessagePath=copied };
            }
            CheckSpace(new FileInfo(source.MessagePath).Length,config);
            var candidate=JsonSerializer.Deserialize<ArchiveItem>(JsonSerializer.Serialize(row))!;
            candidate.StageMsgPath=ArchiveFiles.Child(job.RootDirectory,source.MessagePath); candidate.StageMsgSha256=source.MessageSha256;
            candidate.SnapshotHash=source.Hash; candidate.SourceChangeKey=source.ChangeKey; SaveRelevantOriginal(candidate,source);
            var volume=adopt is null ? Allocate(job,candidate,config) : Volume(candidate);
            candidate.VolumeNumber=volume.Number;
            string manifest=Path.Combine(folder,"repair-intent.json");
            string token=adopt is null ? Guid.NewGuid().ToString("N") : row.Token;
            ArchiveFiles.WriteNew(manifest,new { Version="archive-repair-v1", Previous=previous, Source=source, DestinationToken=token,
                NewVolume=volume.Number, SourceUnknown=sourceUnknown, Adopt=adopt, Utc=DateTimeOffset.UtcNow });
            candidate.RepairManifestPath=manifest; candidate.RepairManifestSha256=ArchiveFiles.Sha(manifest); candidate.RepairToken=token;
            candidate.RepairSourceUnmanaged=sourceUnknown; candidate.RepairPhase=adopt is null ? "ImportStarted" : "Adopted";
            candidate.TargetEntryId=adopt?.EntryId; candidate.TargetStoreId=adopt?.StoreId; candidate.TargetFolderId=adopt?.FolderId;
            candidate.TargetSnapshotHash=null; candidate.TargetFieldsPath=null; candidate.TargetRelevantHash=null; candidate.Error=null;
            ledger.UpdateItem(candidate);
            foreach(var property in typeof(ArchiveItem).GetProperties().Where(x=>x.CanRead && x.CanWrite)) property.SetValue(row,property.GetValue(candidate));
            freshIntent=true; Fault?.Invoke("after-repair-intent",row);
        }
        if (ArchiveFiles.Sha(ArchiveFiles.Child(job.RootDirectory,row.RepairManifestPath!))!=row.RepairManifestSha256)
            throw new ArchiveReviewException("Manifesto de reparo alterado.");
        CheckBackupFile(job,row);
        var part=Volume(row); EnsureVolume(part);
        if (row.RepairPhase == "Adopted" && row.TargetEntryId is not null)
            backend.MarkCopy(part,row,new(row.TargetStoreId!,row.TargetFolderId!,row.TargetEntryId));
        var matches=backend.FindToken(part,row);
        if (matches.Matches.Count>1) throw new ArchiveReviewException("Reparo ambíguo; cópias preservadas.");
        var destination=matches.Matches.SingleOrDefault();
        if (destination is null)
        {
            if (!freshIntent) destination=RecoverUnmarked(job,row,part);
            else destination=backend.Import(part,row);
            Fault?.Invoke("after-repair-copy-before-checkpoint",row);
        }
        row.TargetEntryId=destination.EntryId; row.TargetStoreId=destination.StoreId; row.TargetFolderId=destination.FolderId;
        row.RepairPhase="CopyPresent"; ledger.UpdateItem(row);
        var backup=backend.BackupSnapshot(row,Evidence(job,row,"repair-backup-check"));
        if (!RecordedBackupMatches(job,row,backup)) throw new ArchiveReviewException("Backup do reparo não corresponde ao registro.");
        var copy=backend.TargetSnapshot(part,row,destination,Evidence(job,row,"repair-target-check"));
        if (!RelevantMessageValidation.Equivalent(backup,copy))
            throw new ArchiveReviewException("Nova cópia ainda diverge: "+RelevantMessageValidation.DifferenceSummary(backup,copy));
        row.TargetSnapshotHash=copy.Hash; row.TargetFieldsPath=Path.Combine(Path.GetDirectoryName(copy.MessagePath)!,"field-hashes.json"); SaveRelevantTarget(row,copy);
        part.ObservedBytes=backend.MeasureVolume(part);
        if(part.ObservedBytes>config.HardLimitBytes-config.ReserveBytes) throw new ArchiveReviewException("Limite seguro do PST atingido.");
        ledger.UpdateVolume(part); row.State=ArchiveItemState.Copied; row.Error=null; row.RepairPhase="Compared"; ledger.UpdateItem(row);
    }

    private string RepairDirectory(ArchiveJob job,ArchiveItem row,string stage)
    { string path=Evidence(job,row,stage); Directory.CreateDirectory(path); return path; }

    private void ConfirmBackupOnly(ArchiveJob job,ArchiveItem row,ArchiveVolume volume,DestinationIdentity identity,bool allowRetained,SourceLocation? observed)
    {
        CheckBackupFile(job,row);
        var disk=backend.ReopenVolume(volume);
        if(!disk.Reopened || !disk.ExclusiveDiskRead && !allowRetained) throw new ArchiveReviewException("Gravação do PST não confirmada.");
        var backup=backend.BackupSnapshot(row,Evidence(job,row,"absent-backup-check"));
        var target=backend.TargetSnapshot(volume,row,identity,Evidence(job,row,"absent-target-check"));
        if(!RecordedBackupMatches(job,row,backup) || !RelevantMessageValidation.Equivalent(backup,target))
            throw new ArchiveReviewException("Cópia no PST não equivale ao backup preservado.");
        ArchiveFiles.WriteNew(Path.Combine(RepairDirectory(job,row,"absent-source-proof"),"proof.json"),new
        { Key=row.Key, BackupSha=row.StageMsgSha256, Destination=identity, TargetHash=target.Hash, disk, SourceNotLocated=observed is null, ObservedFolder=observed?.FolderId, OriginalDeletionClaimed=false, Utc=DateTimeOffset.UtcNow });
        row.TargetEntryId=identity.EntryId; row.TargetStoreId=identity.StoreId; row.TargetFolderId=identity.FolderId;
        row.TargetSnapshotHash=target.Hash; row.TargetFieldsPath=Path.Combine(Path.GetDirectoryName(target.MessagePath)!,"field-hashes.json"); SaveRelevantTarget(row,target);
        row.State=ArchiveItemState.BackupVerifiedSourceUnmanaged; row.Error=null; ledger.UpdateItem(row);
    }
}
