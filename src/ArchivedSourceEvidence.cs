internal static class ArchivedSourceEvidence
{
    internal static bool Matches(ArchiveItem item, ArchiveSnapshot current, ArchiveSnapshot? replay = null)
    {
        string message = item.StageMsgPath ?? throw new ArchiveReviewException("Backup original ausente.");
        if (ArchiveFiles.Sha(message) != item.StageMsgSha256) throw new ArchiveReviewException("Backup original alterado.");
        string directory = Path.GetDirectoryName(message)!;
        var fields = ArchiveFiles.ReadFields(Path.Combine(directory, "field-hashes.json"));
        if (ArchiveFiles.HashFields(fields) != item.SnapshotHash) throw new ArchiveReviewException("Evidência original alterada.");
        Dictionary<string, string>? relevant = null;
        if (item.SourceRelevantHash is not null)
        {
            relevant = ArchiveFiles.ReadFields(Path.Combine(directory, "relevant-fields.json"));
            if (ArchiveFiles.HashFields(relevant) != item.SourceRelevantHash) throw new ArchiveReviewException("Evidência das informações relevantes alterada.");
        }
        var original = DeliveryReportEvidence.BindSubjectProof(new ArchiveSnapshot(fields, message, item.StageMsgSha256!, item.SourceChangeKey, relevant));
        if(item.RepairToken is not null)original=RelevantMessageValidation.RecoveryProfile(original);
        if (replay is not null)
        {
            if (!EmbeddedMessageEvidence.CanReplay(original, replay)) throw new ArchiveReviewException("Replay do e-mail anexado não corresponde ao backup original.");
            original = original with { RelevantFields = replay.RelevantFields };
        }
        return RelevantMessageValidation.Equivalent(original, current);
    }
}
