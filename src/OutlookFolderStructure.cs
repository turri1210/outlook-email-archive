using System.Runtime.InteropServices;
using System.Text.Json;

internal sealed partial class OutlookArchiveBackend
{
    private readonly Dictionary<int, (string Id, string Name)> sourceDefaults = new();
    internal object FolderStructureSmoke(string directory, string sourceStoreId)
    {
        directory = Path.GetFullPath(directory); ArchiveFiles.ValidateLocalDirectory(directory);
        if (Directory.Exists(directory)) throw new IOException("Use um diretório de teste novo.");
        Directory.CreateDirectory(directory); string pst = Path.Combine(directory, "folder-structure-test.pst");
        object? store = null, root = null, inbox = null, subfolder = null, control = null, source = null;
        try
        {
            InitializeFolderStructure(sourceStoreId);
            if (!sourceDefaults.TryGetValue(6, out var sourceInbox)) throw new ArchiveReviewException("Caixa de Entrada padrão da origem indisponível.");
            var config = new ArchiveConfig { SourceStoreId = sourceStoreId, PreserveFolderStructure = true };
            source = ((dynamic)session!).GetFolderFromID(sourceInbox.Id, sourceStoreId);
            RegisterFolderDestination(source, [sourceInbox.Name], config);
            if (config.FolderMappings[JsonSerializer.Serialize(new[] { sourceInbox.Name })].DefaultFolderType != 6) throw new Exception("Inbox não identificado pelo tipo.");
            config.FolderMappings[JsonSerializer.Serialize(new[] { sourceInbox.Name, "AC BR" })] = new(6, sourceInbox.Name, ["AC BR"]);
            ((dynamic)session!).AddStoreEx(pst, 2); store = FindStore(pst)!; root = ((dynamic)store).GetRootFolder();
            inbox = NamedFolder(root, ["Inbox"], true);
            Console.WriteLine($"Teste: pasta criada {(string)((dynamic)inbox!).Name}; id {(string)((dynamic)inbox!).EntryID}; caminho {(string)((dynamic)inbox!).FolderPath}");
            Release(inbox); inbox = null;
            var volume = new ArchiveVolume { JobId = Guid.NewGuid().ToString("N"), Number = 1, PstPath = pst, IsExisting = true,
                SourceStoreId = sourceStoreId, PreserveFolderStructure = true, FolderMappings = config.FolderMappings };
            CreateVolume(volume);
            inbox = TargetFolder(volume, [sourceInbox.Name], true);
            Console.WriteLine($"Teste: pasta resolvida {(string)((dynamic)inbox!).Name}; id {(string)((dynamic)inbox!).EntryID}");
            subfolder = TargetFolder(volume, [sourceInbox.Name, "AC BR"], true);
            string inboxId = (string)((dynamic)inbox!).EntryID;
            object? parent = null;
            try { parent = ((dynamic)subfolder!).Parent; if (!SameIdentity((string)((dynamic)parent).EntryID, inboxId)) throw new Exception("Subpasta fora do Inbox."); }
            finally { Release(parent); }
            if ((string)((dynamic)inbox!).Name != "Inbox" || (string)((dynamic)subfolder!).Name != "AC BR") throw new Exception("Nomes de destino incorretos.");
            control = OwnedScope(store, volume, false);
            bool hidden = Property(control, "http://schemas.microsoft.com/mapi/proptag/0x10F4000B") is true;
            if (!hidden) throw new Exception("Controle técnico não ficou oculto.");
            return new { Passed = true, SourceDefaultFolder = sourceInbox.Name, Destination = (string)((dynamic)subfolder!).FolderPath,
                StandardFolderType = 6, ControlFolderHidden = hidden, RealMessagesCopied = 0, PstPath = pst };
        }
        finally
        {
            Release(source); Release(control); Release(subfolder); Release(inbox);
            if (root is not null) { try { ((dynamic)session!).RemoveStore(root); } catch (COMException) { } }
            Release(root); Release(store);
        }
    }
    private string? sourceRootFolderId;
    private void InitializeFolderStructure(string storeId)
    {
        sourceDefaults.Clear(); object? store = null, root = null;
        try
        {
            store = ((dynamic)session!).GetStoreFromID(storeId); root = ((dynamic)store).GetRootFolder();
            sourceRootFolderId = (string)((dynamic)root).EntryID;
            foreach (int type in new[] { 3, 4, 5, 6, 16, 23 })
            {
                object? folder = null;
                try { folder = ((dynamic)store).GetDefaultFolder(type); if (folder is not null && !string.IsNullOrWhiteSpace((string)((dynamic)folder).Name)) sourceDefaults[type] = ((string)((dynamic)folder).EntryID, (string)((dynamic)folder).Name); }
                catch (COMException ex) when (ex.HResult is unchecked((int)0x8004010F) or unchecked((int)0x80020009)) { }
                finally { Release(folder); }
            }
        }
        finally { Release(root); Release(store); }
    }
    private void RegisterFolderDestination(object folder, string[] relativePath, ArchiveConfig config)
    {
        object? current = null; var names = new List<string>();
        try
        {
            current = ((dynamic)session!).GetFolderFromID((string)((dynamic)folder).EntryID, config.SourceStoreId);
            for (int depth = 0; depth < 100; depth++)
            {
                string id = (string)((dynamic)current).EntryID;
                if (SameIdentity(id, sourceRootFolderId))
                {
                    names.Reverse(); config.FolderMappings[JsonSerializer.Serialize(relativePath)] = new(null, "", names.ToArray()); return;
                }
                foreach (var standard in sourceDefaults)
                    if (SameIdentity(id, standard.Value.Id))
                    {
                        names.Reverse(); config.FolderMappings[JsonSerializer.Serialize(relativePath)] = new(standard.Key, standard.Value.Name, names.ToArray()); return;
                    }
                names.Add((string)((dynamic)current).Name);
                object parent = ((dynamic)current).Parent; Release(current); current = parent;
            }
            throw new ArchiveReviewException("Hierarquia da origem excedeu o limite seguro.");
        }
        finally { Release(current); }
    }
    private object? StructuredTargetFolder(object store, ArchiveVolume volume, string[] relativePath, bool create)
    {
        if (!volume.FolderMappings.TryGetValue(JsonSerializer.Serialize(relativePath), out var mapping))
            throw new ArchiveReviewException("Pasta sem mapeamento congelado; arquivamento bloqueado.");
        object? current = null;
        try
        {
            if (mapping.DefaultFolderType is int type)
            {
                try { current = ((dynamic)store).GetDefaultFolder(type); }
                catch (COMException ex) when (ex.HResult is unchecked((int)0x8004010F) or unchecked((int)0x80020009)) { }
                // A new PST can return an unnamed internal placeholder instead
                // of an actual Inbox/Sent Items. Do not archive into that folder.
                if (current is not null && string.IsNullOrWhiteSpace((string)((dynamic)current).Name)) { Release(current); current = null; }
                if (current is null)
                {
                    object? root = null;
                    try
                    {
                        root = ((dynamic)store).GetRootFolder();
                        string[] aliases = FolderStructureRules.StandardNames(type, mapping.DefaultFolderName);
                        current = NamedFolder(root, aliases, false);
                        if (current is null && create) current = NamedFolder(root, [mapping.DefaultFolderName], true);
                    }
                    finally { Release(root); }
                }
            }
            else current = ((dynamic)store).GetRootFolder();
            if (current is null) return null;
            if (!SameIdentity((string)((dynamic)current).StoreID, volume.StoreId)) throw new ArchiveReviewException("Pasta padrão fora do PST registrado.");
            foreach (string name in mapping.Subfolders)
            {
                object? next = NamedFolder(current, [name], create);
                Release(current); current = next;
                if (current is null) return null;
            }
            object result = current; current = null; return result;
        }
        finally { Release(current); }
    }
    private static object? NamedFolder(object parent, string[] names, bool create)
    {
        object? folders = null, found = null;
        try
        {
            folders = ((dynamic)parent).Folders;
            for (int i = 1; i <= (int)((dynamic)folders).Count; i++)
            {
                object? candidate = ((dynamic)folders)[i];
                if (names.Contains((string)((dynamic)candidate).Name, StringComparer.OrdinalIgnoreCase))
                {
                    if (found is not null) { Release(candidate); throw new ArchiveReviewException("Há pastas equivalentes ambíguas no PST; confira os nomes antes de continuar."); }
                    found = candidate;
                }
                else Release(candidate);
            }
            if (found is null && create) found = ((dynamic)folders).Add(names[0], 6);
            object? result = found; found = null; return result;
        }
        finally { Release(found); Release(folders); }
    }
}

internal static class FolderStructureRules
{
    internal static bool StoreIdContainsPath(string id, string path)
    {
        byte[] bytes;
        try { bytes = Convert.FromHexString(id); } catch (FormatException) { return false; }
        string expected = Path.GetFullPath(path) + '\0'; int length = expected.Length * 2;
        for (int offset = 0; offset + length <= bytes.Length; offset++)
            if (string.Equals(System.Text.Encoding.Unicode.GetString(bytes, offset, length), expected, StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }
    internal static string[] StandardNames(int type, string sourceName) => (type switch
    {
        6 => new[] { sourceName, "Inbox", "Caixa de Entrada" },
        5 => new[] { sourceName, "Sent Items", "Itens Enviados" },
        3 => new[] { sourceName, "Deleted Items", "Itens Excluídos" },
        16 => new[] { sourceName, "Drafts", "Rascunhos" },
        4 => new[] { sourceName, "Outbox", "Caixa de Saída" },
        23 => new[] { sourceName, "Junk Email", "Lixo Eletrônico" },
        _ => throw new ArchiveReviewException("Tipo de pasta padrão não suportado.")
    }).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
}
