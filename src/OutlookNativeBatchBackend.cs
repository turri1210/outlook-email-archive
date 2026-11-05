using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

internal sealed partial class OutlookArchiveBackend : IArchiveBatchBackend
{
    private const string BatchOwnerProperty = "http://schemas.microsoft.com/mapi/string/{DDB565AD-3206-4E72-B044-C54157137867}/BatchOwner";
    private const string BatchContainerName = "Lotes temporários de arquivamento";
    private readonly Dictionary<string, TokenIndex> batchRecoveryIndexes = new(StringComparer.Ordinal);
    private static string NativeCopyPath => Path.Combine(AppContext.BaseDirectory, "native-copy.exe");

    public bool CanBatch(ArchiveItem item)
    {
        // Automatic-reply templates use the conservative single-message path.
        if(item.SourceMessageClass==DeliveryReportEvidence.AutomaticReplyClass)return false;
        if (item.SourceMessageClass == DeliveryReportEvidence.MessageClass)
        {
            if(!Environment.Is64BitProcess || !File.Exists(NativeCopyPath))return false;
            object? report=null;
            try
            {
                report=OpenSource(item,item.SourceEntryId,item.SourceFolderId,item.SourceRecordKey);
                return !BatchReceiptValue(Property(report,"http://schemas.microsoft.com/mapi/proptag/0x0029000B")) && !BatchReceiptValue(Property(report,"http://schemas.microsoft.com/mapi/proptag/0x0C06000B"));
            }
            finally{Release(report);}
        }
        if (bulkSourceIdentities.TryGetValue(SourceBulkKey(item), out var cached))
        { _ = RequireBulkIdentity(cached); return cached.CanBatch; }
        if (!Environment.Is64BitProcess || !File.Exists(NativeCopyPath)) return false;
        object? source = null;
        try
        {
            source = OpenFastSource(item, ["0x0029000B", "0x0C06000B"], out var values);
            if (values[3] is not int flags || flags < 0)
                throw new ArchiveReviewException("Sinalizadores da origem indisponíveis para cópia em lote.");
            return (flags & 0x300) == 0 && !BatchReceiptValue(values[FastGuardTags.Length]) && !BatchReceiptValue(values[FastGuardTags.Length + 1]);
        }
        finally { Release(source); }
    }

    private static bool BatchReceiptValue(object? value)
    {
        if (FastMissing(value)) return false;
        return value is bool requested ? requested : throw new ArchiveReviewException("Propriedade de recibo inválida; cópia nativa bloqueada.");
    }

    private static void ValidateBatch(ArchiveVolume volume, ArchiveBatch batch)
    {
        if (!Guid.TryParseExact(batch.Id, "N", out _) || !Guid.TryParseExact(batch.JobId, "N", out _) ||
            batch.JobId != volume.JobId || batch.VolumeNumber != volume.Number || batch.Items.Count is < 1 or > 100 ||
            batch.Items.Distinct(StringComparer.Ordinal).Count() != batch.Items.Count)
            throw new ArchiveReviewException("Identidade ou escopo do lote inválido.");
        string directory = Path.GetFullPath(batch.DirectoryPath).TrimEnd(Path.DirectorySeparatorChar);
        if (Path.GetFileName(directory) != batch.Id || Path.GetFileName(Path.GetDirectoryName(directory)) != "batches" ||
            Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(directory))) != batch.JobId)
            throw new ArchiveReviewException("Pasta de evidências fora da execução do lote.");
        ArchiveFiles.ValidateLocalDirectory(directory);
        if (!Directory.Exists(directory)) throw new ArchiveReviewException("Pasta de evidências do lote ausente.");
    }

    public void AssertBatchIdle(ArchiveBatch batch)
    {
        // A worker survives an interrupted engine, so both process-start and
        // native-operation leases must be free before any reconciliation.
        using var launch = new FileStream(Path.Combine(batch.DirectoryPath, "launch.lock"), FileMode.OpenOrCreate,
            FileAccess.ReadWrite, FileShare.None);
        using var worker = new FileStream(Path.Combine(batch.DirectoryPath, "worker.lock"), FileMode.OpenOrCreate,
            FileAccess.ReadWrite, FileShare.None);
    }

    public BatchFolder PrepareBatch(ArchiveVolume volume, ArchiveBatch batch)
    {
        ValidateBatch(volume, batch); AttachVolume(volume); AssertBatchIdle(batch);
        object? store = null, scope = null, container = null, folder = null;
        try
        {
            store = FindStore(volume.PstPath, volume.StoreId) ?? throw new ArchiveReviewException("PST do lote indisponível.");
            scope = OwnedScope(store, volume, false);
            container = OwnedBatchChild(scope, BatchContainerName, Owner(volume) + ":batches", true);
            folder = OwnedBatchChild(container, "Lote-" + batch.Id, BatchOwner(volume, batch), true);
            RequireNoBatchChildren(folder);
            string id = (string)((dynamic)folder).EntryID;
            if (batch.FolderId is not null && !BatchSame(batch.FolderId, id))
                throw new ArchiveReviewException("A pasta do lote foi substituída.");
            if (batch.StoreId is not null && !BatchSame(batch.StoreId, volume.StoreId!))
                throw new ArchiveReviewException("O PST do lote foi substituído.");
            if (batch.FolderId is null && BatchCount(folder) != 0)
                throw new ArchiveReviewException("Pasta de lote não registrada contém mensagens; revisão necessária.");
            SetBatchHidden(container, false); SetBatchHidden(folder, false);
            return new(volume.StoreId!, id);
        }
        finally { Release(folder); Release(container); Release(scope); Release(store); }
    }

    public void CopyBatch(ArchiveVolume volume, ArchiveBatch batch, IReadOnlyList<ArchiveItem> items)
    {
        ValidateBatch(volume, batch); AttachVolume(volume);
        if (items.Count != batch.Items.Count || !items.Select(x => x.Key).SequenceEqual(batch.Items) ||
            items.Any(x => x.JobId != batch.JobId || !BatchSame(x.SourceStoreId, batch.SourceStoreId) ||
                !BatchSame(x.SourceFolderId, batch.SourceFolderId) || !x.RelativePath.SequenceEqual(batch.RelativePath)) ||
            BatchSame(batch.SourceStoreId, volume.StoreId!))
            throw new ArchiveReviewException("Itens fora da pasta, volume ou escopo congelado do lote.");
        using var launch = new FileStream(Path.Combine(batch.DirectoryPath, "launch.lock"), FileMode.OpenOrCreate,
            FileAccess.ReadWrite, FileShare.None);
        using (var idle = new FileStream(Path.Combine(batch.DirectoryPath, "worker.lock"), FileMode.OpenOrCreate,
            FileAccess.ReadWrite, FileShare.None)) { }
        object? folder = null;
        try
        {
            folder = OpenBatchFolder(volume, batch);
            if (BatchCount(folder) != 0) throw new ArchiveReviewException("O lote já possui resultados; repetição da cópia bloqueada.");
            if (batch.FastCopy && !HasFreshBulkScope(items)) throw new ArchiveReviewException("Escopo nativo rápido não foi identificado integralmente antes da cópia.");
            foreach (var row in batch.FastCopy ? Array.Empty<ArchiveItem>() : items)
            {
                if (!CanBatch(row)) throw new ArchiveReviewException("A condição de leitura ou recibo mudou; lote não enviado ao Outlook.");
                if (!string.IsNullOrEmpty(row.SourceChangeKey) && !SourceMatchesChangeKey(row, row.SourceChangeKey))
                    throw new ArchiveReviewException("A origem mudou antes da cópia em lote.");
            }
            string handoff = Path.Combine(batch.DirectoryPath, "handoff-" + Guid.NewGuid().ToString("N") + ".txt");
            using (var stream = new FileStream(handoff, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                using var writer = new StreamWriter(stream, new UTF8Encoding(false), leaveOpen: true);
                foreach (string line in new[] { "production-copy-only-v1", Path.Combine(batch.DirectoryPath, "worker.lock"),
                    batch.SourceStoreId, batch.SourceFolderId, volume.StoreId!, batch.FolderId! }.Concat(items.Select(x => x.SourceEntryId)))
                    writer.WriteLine(line);
                writer.Flush(); stream.Flush(true);
            }
            using var process = new Process { StartInfo = new ProcessStartInfo(NativeCopyPath)
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true } };
            process.StartInfo.ArgumentList.Add(handoff);
            if (!process.Start()) throw new IOException("Não foi possível iniciar a cópia nativa.");
            var output = process.StandardOutput.ReadToEndAsync(); var errors = process.StandardError.ReadToEndAsync();
            process.WaitForExit();
            string stdout = output.GetAwaiter().GetResult(), stderr = errors.GetAwaiter().GetResult();
            File.WriteAllText(Path.Combine(batch.DirectoryPath, "worker-output.json"), stdout);
            File.WriteAllText(Path.Combine(batch.DirectoryPath, "worker-errors.txt"), stderr);
            tokenIndexes.Remove(volume.PstPath); batchRecoveryIndexes.Remove(batch.Id);
            if (process.ExitCode != 0) throw new ArchiveReviewException("O Outlook retornou um resultado parcial ou uma falha. O lote será reconciliado sem repetir a cópia.");
            using var result = JsonDocument.Parse(stdout);
            var value = result.RootElement;
            if (value.GetProperty("Stage").GetString() != "copy-returned" || value.GetProperty("HResult").GetString() != "0x0" ||
                value.GetProperty("Requested").GetInt32() != items.Count || value.GetProperty("MoveFlag").GetBoolean())
                throw new ArchiveReviewException("O componente nativo não confirmou o escopo exato da cópia.");
        }
        finally { Release(folder); }
    }

    public List<DestinationIdentity> BatchCandidates(ArchiveVolume volume, ArchiveBatch batch)
    {
        ValidateBatch(volume, batch); AttachVolume(volume); AssertBatchIdle(batch);
        if (batch.FastCopy) return ReadBulkBatchCandidates(volume, batch);
        object? folder = null, items = null; var result = new List<DestinationIdentity>();
        try
        {
            folder = OpenBatchFolder(volume, batch); int count = BatchCount(folder);
            if (count > batch.Items.Count) throw new ArchiveReviewException("O lote contém cópias extras; adoção automática bloqueada.");
            items = ((dynamic)folder).Items;
            for (int i = 1; i <= count; i++)
            {
                object? item = null, parent = null;
                try
                {
                    item = ((dynamic)items)[i]; parent = ((dynamic)item).Parent;
                    if (!BatchSame((string)((dynamic)parent).StoreID, volume.StoreId!) || !BatchSame((string)((dynamic)parent).EntryID, batch.FolderId!))
                        throw new ArchiveReviewException("Candidato fora da pasta própria do lote.");
                    result.Add(new(volume.StoreId!, batch.FolderId!, (string)((dynamic)item).EntryID));
                }
                finally { Release(parent); Release(item); }
            }
            if (BatchCount(folder) != count || result.Select(x => x.EntryId).Distinct(StringComparer.OrdinalIgnoreCase).Count() != count)
                throw new ArchiveReviewException("O conteúdo do lote mudou durante a conferência.");
            return result;
        }
        finally { Release(items); Release(folder); }
    }

    public void MarkBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem row, DestinationIdentity identity)
    {
        ValidateBatch(volume, batch); AssertBatchIdle(batch); AttachVolume(volume);
        object? copy = null;
        try
        {
            copy = LocateBatchCopy(volume, batch, row, identity, requireToken: false);
            string? old = Property(copy, TokenProperty) as string;
            if (!string.IsNullOrEmpty(old) && old != row.Token) throw new ArchiveReviewException("A cópia possui outro identificador.");
            if (old != row.Token) { SetProperty(copy, TokenProperty, row.Token); ((dynamic)copy).Save(); }
            if ((string?)Property(copy, TokenProperty) != row.Token) throw new ArchiveReviewException("O identificador da cópia não persistiu.");
            tokenIndexes.Remove(volume.PstPath);
        }
        finally { Release(copy); }
    }

    public DestinationIdentity FinalizeBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem row, DestinationIdentity identity)
    {
        ValidateBatch(volume, batch); AssertBatchIdle(batch); AttachVolume(volume);
        object? copy = null, parent = null, final = null, moved = null;
        try
        {
            copy = LocateBatchCopy(volume, batch, row, identity, requireToken: true);
            parent = ((dynamic)copy).Parent; final = TargetFolder(volume, row.RelativePath, true)!;
            if (!BatchSame((string)((dynamic)final).StoreID, volume.StoreId!)) throw new ArchiveReviewException("Pasta final fora do PST do lote.");
            string finalId = (string)((dynamic)final).EntryID;
            if (BatchSame((string)((dynamic)parent).EntryID, finalId)) moved = copy;
            else moved = ((dynamic)copy).Move(final); // Stored copy only, inside the SAME PST.
            var result = new DestinationIdentity(volume.StoreId!, finalId, (string)((dynamic)moved).EntryID);
            GuardBatchCopy(volume, batch, row, moved, allowStaging: false);
            if ((string?)Property(moved, TokenProperty) != row.Token) throw new ArchiveReviewException("A movimentação interna perdeu a identificação.");
            tokenIndexes.Remove(volume.PstPath);
            if (batchRecoveryIndexes.TryGetValue(batch.Id, out var index)) index.Tokens[row.Token] = [result];
            return result;
        }
        finally { if (!ReferenceEquals(moved, copy)) Release(moved); Release(final); Release(parent); Release(copy); }
    }

    public void CloseBatch(ArchiveVolume volume, ArchiveBatch batch)
    {
        ValidateBatch(volume, batch); AssertBatchIdle(batch); AttachVolume(volume);
        object? folder = null;
        try
        {
            folder = OpenBatchFolder(volume, batch);
            if (BatchCount(folder) != 0) throw new ArchiveReviewException("A pasta do lote ainda contém mensagens e será preservada.");
            // No folder deletion: Outlook would move it to Deleted Items. Hide
            // only this empty, owned staging folder; never hide partial results.
            SetBatchHidden(folder, true);
            batchRecoveryIndexes.Remove(batch.Id); tokenIndexes.Remove(volume.PstPath);
        }
        finally { Release(folder); }
    }

    private object LocateBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem row, DestinationIdentity identity, bool requireToken)
    {
        if (!batch.Items.Contains(row.Key) || identity.StoreId != volume.StoreId) throw new ArchiveReviewException("A cópia não pertence a este lote.");
        object? copy = null;
        try
        {
            copy = ((dynamic)session!).GetItemFromID(identity.EntryId, volume.StoreId);
            GuardBatchCopy(volume, batch, row, copy!, allowStaging: true);
            string? token = Property(copy!, TokenProperty) as string;
            if (requireToken && token != row.Token) throw new ArchiveReviewException("Cópia sem o identificador esperado.");
            if (!string.IsNullOrEmpty(token) && token != row.Token) throw new ArchiveReviewException("Outro item reivindicou a cópia.");
            object result = copy!; copy = null; return result;
        }
        catch (COMException ex) when (ex.HResult is unchecked((int)0x8004010F) or unchecked((int)0x80040107))
        { /* EntryID can change after an interrupted within-PST move. Recover by token. */ }
        finally { Release(copy); }
        if (!batchRecoveryIndexes.TryGetValue(batch.Id, out var index))
        {
            index = new TokenIndex(); object? staging = null, final = null;
            try
            {
                staging = OpenBatchFolder(volume, batch); ScanTokens(staging, volume.StoreId!, index);
                final = TargetFolder(volume, row.RelativePath, false);
                if (final is not null) ScanTokens(final, volume.StoreId!, index);
            }
            finally { Release(final); Release(staging); }
            batchRecoveryIndexes[batch.Id] = index;
        }
        var matches = index.Tokens.GetValueOrDefault(row.Token) ?? [];
        if (matches.Count != 1) throw new ArchiveReviewException("A movimentação interrompida não possui uma cópia identificada única.");
        copy = ((dynamic)session!).GetItemFromID(matches[0].EntryId, volume.StoreId);
        try { GuardBatchCopy(volume, batch, row, copy!, allowStaging: true); object result = copy!; copy = null; return result; }
        finally { Release(copy); }
    }

    private void GuardBatchCopy(ArchiveVolume volume, ArchiveBatch batch, ArchiveItem row, object copy, bool allowStaging)
    {
        object? parent = null, final = null;
        try
        {
            parent = ((dynamic)copy).Parent;
            if (!BatchSame((string)((dynamic)parent).StoreID, volume.StoreId!) || !(row.SourceMessageClass == DeliveryReportEvidence.MessageClass ? DeliveryReportEvidence.IsReport(copy) : (int)((dynamic)copy).Class == 43 && (string)((dynamic)copy).MessageClass == "IPM.Note"))
                throw new ArchiveReviewException("Gravação bloqueada fora da cópia no PST.");
            string folder = (string)((dynamic)parent).EntryID;
            if (allowStaging && BatchSame(folder, batch.FolderId!))
            {
                if ((string?)Property(parent, BatchOwnerProperty) != BatchOwner(volume, batch)) throw new ArchiveReviewException("Pasta temporária sem proprietário.");
                return;
            }
            final = TargetFolder(volume, row.RelativePath, false);
            if (final is null || !BatchSame(folder, (string)((dynamic)final).EntryID))
                throw new ArchiveReviewException("A cópia foi movida para fora da pasta planejada.");
        }
        finally { Release(final); Release(parent); }
    }

    private object OpenBatchFolder(ArchiveVolume volume, ArchiveBatch batch)
    {
        if (batch.StoreId != volume.StoreId || string.IsNullOrEmpty(batch.FolderId)) throw new ArchiveReviewException("Pasta do lote não registrada.");
        object? store = null, scope = null, container = null, folder = null;
        try
        {
            store = FindStore(volume.PstPath, volume.StoreId)!; scope = OwnedScope(store, volume, false);
            container = OwnedBatchChild(scope, BatchContainerName, Owner(volume) + ":batches", false);
            folder = OwnedBatchChild(container, "Lote-" + batch.Id, BatchOwner(volume, batch), false);
            RequireNoBatchChildren(folder);
            if (!BatchSame((string)((dynamic)folder).StoreID, volume.StoreId!) || !BatchSame((string)((dynamic)folder).EntryID, batch.FolderId))
                throw new ArchiveReviewException("Identidade da pasta própria do lote não confere.");
            object result = folder; folder = null; return result;
        }
        finally { Release(folder); Release(container); Release(scope); Release(store); }
    }

    private static object OwnedBatchChild(object parent, string name, string owner, bool create)
    {
        object? folders = null, found = null;
        try
        {
            folders = ((dynamic)parent).Folders;
            for (int i = 1; i <= (int)((dynamic)folders).Count; i++)
            {
                object? candidate = ((dynamic)folders)[i];
                if ((string)((dynamic)candidate).Name == name)
                {
                    if (found is not null) { Release(candidate); throw new ArchiveReviewException("Pasta de lote ambígua."); }
                    found = candidate;
                }
                else Release(candidate);
            }
            if (found is null && create) { found = ((dynamic)folders).Add(name, 6); SetProperty(found, BatchOwnerProperty, owner); }
            if (found is null || (string?)Property(found, BatchOwnerProperty) != owner) throw new ArchiveReviewException("Pasta de lote ausente ou sem propriedade válida.");
            object result = found; found = null; return result;
        }
        finally { Release(found); Release(folders); }
    }

    private static int BatchCount(object folder)
    { object? items = null; try { items = ((dynamic)folder).Items; return (int)((dynamic)items).Count; } finally { Release(items); } }
    private static void RequireNoBatchChildren(object folder)
    { object? children = null; try { children = ((dynamic)folder).Folders; if ((int)((dynamic)children).Count != 0) throw new ArchiveReviewException("A pasta do lote contém subpastas inesperadas."); } finally { Release(children); } }
    private bool BatchSame(string left, string right) => OutlookIdentity.AreSame(left, right,
        (a, b) => (bool)((dynamic)session!).CompareEntryIDs(a, b));
    private static string BatchOwner(ArchiveVolume volume, ArchiveBatch batch) => Owner(volume) + ":" + batch.Id;
    private static void SetBatchHidden(object folder, bool hidden)
    {
        object? accessor = null;
        try { accessor = ((dynamic)folder).PropertyAccessor; ((dynamic)accessor).SetProperty("http://schemas.microsoft.com/mapi/proptag/0x10F4000B", hidden); }
        catch (COMException) { /* Cosmetic only; ownership and copy guards remain required. */ }
        finally { Release(accessor); }
    }
}
