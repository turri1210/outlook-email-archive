using System.Buffers.Binary;
using System.ComponentModel;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

// Creates synthetic messages ONLY inside a new, explicitly selected local PST.
// No default CreateItem, Send, recipient resolution, real-mail lookup, or deletion.
internal static class LocalFixtureCheck
{
    internal static int Probe(string directory)
    {
        object? app = null, session = null, folder = null, items = null, item = null, parent = null;
        try
        {
            var manifest = JsonSerializer.Deserialize<FixtureManifest>(File.ReadAllText(Path.Combine(directory, "manifest.json")))
                ?? throw new InvalidDataException("Manifesto inválido.");
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            folder = ((dynamic)session).GetFolderFromID(manifest.SourceFolderId, manifest.SourceStoreId);
            items = ((dynamic)folder).Items;
            int before = (int)((dynamic)items).Count;
            item = ((dynamic)items).Add(0);
            parent = ((dynamic)item).Parent;
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                StoreMatches = (string)((dynamic)parent).StoreID == manifest.SourceStoreId,
                FolderMatches = (string)((dynamic)parent).EntryID == manifest.SourceFolderId,
                ParentFolderName = (string)((dynamic)parent).Name,
                Class = (int)((dynamic)item).Class,
                MessageClass = (string)((dynamic)item).MessageClass,
                Saved = (bool)((dynamic)item).Saved,
                HasEntryId = !string.IsNullOrEmpty((string?)((dynamic)item).EntryID),
                CountBefore = before, CountAfter = (int)((dynamic)items).Count
            }));
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine("Diagnóstico: " + ex.GetType().Name + ": " + ex.Message); return 1; }
        finally { Release(parent); Release(item); Release(items); Release(folder); Release(session); Release(app); }
    }

    private const int FixtureCount = 5;
    private const string FixtureAddress = "fixture@example.invalid";

    internal static int Execute(string[] args)
    {
        if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0]))
        { Console.Error.WriteLine("Uso: local-fixture-check <novo-diretorio-absoluto-fora-do-OneDrive>"); return 2; }
        DurableJournal? journal = null;
        string? directory = null;
        bool directoryCreated = false;
        var manifest = new FixtureManifest();
        try
        {
            if (!Path.IsPathFullyQualified(args[0])) throw new ArgumentException("Use um caminho absoluto novo.");
            directory = Path.GetFullPath(args[0]).TrimEnd(Path.DirectorySeparatorChar);
            string parent = Path.GetDirectoryName(directory) ?? throw new ArgumentException("Diretório pai inválido.");
            if (new DirectoryInfo(directory).Name.Contains("OneDrive", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("O diretório de trabalho não pode ficar no OneDrive.");
            // Validate the proposed directory as a not-yet-existing run path; both
            // parents already exist, so nothing must be created just to validate.
            string probePst = Path.Combine(parent, ".fixture-validation-" + Guid.NewGuid().ToString("N") + ".pst");
            PilotCommand.ValidateDestination(probePst, directory, FixtureCount);
            CreateDirectoryExclusive(directory);
            directoryCreated = true;
            string sourcePst = Path.Combine(directory, "source-fixtures.pst");
            string targetPst = Path.Combine(directory, "destination-fixtures.pst");
            string pilotRun = Path.Combine(directory, "pilot-run");
            PilotCommand.ValidateDestination(sourcePst, pilotRun, FixtureCount);
            PilotCommand.ValidateDestination(targetPst, pilotRun, FixtureCount);
            string inputs = Path.Combine(directory, "inputs");
            CreateDirectoryExclusive(inputs);
            journal = new DurableJournal(Path.Combine(directory, "fixture-journal.jsonl"));
            manifest.SourcePstPath = sourcePst;
            manifest.TargetPstPath = targetPst;
            manifest.PilotRunDirectory = pilotRun;
            journal.Write(new { State = "fixture-set-planned", Count = FixtureCount, SourcePst = sourcePst, TargetPst = targetPst });

            WriteInput(inputs, "anexo-ação.txt", new UTF8Encoding(false).GetBytes("Arquivo artificial UTF-8: ação, café, coração, € e 漢字.\r\n"), manifest);
            WriteInput(inputs, "dados-binários-64KB.bin", PatternBytes(64 * 1024), manifest);
            WriteInput(inputs, "imagem-ícone.png", CreatePng(), manifest);
            WriteInput(inputs, "relatório-RTF.txt", new UTF8Encoding(false).GetBytes("Anexo artificial do caso RTF: certificação e informação.\r\n"), manifest);
            WriteInput(inputs, "amostra-binária-2MB.bin", PatternBytes(2 * 1024 * 1024), manifest);

            CreateFixtureMessages(sourcePst, inputs, journal, manifest);
            manifest.Status = "local-source-fixtures-created";
            WriteJsonNew(Path.Combine(directory, "manifest.json"), manifest);
            journal.Write(new { State = "local-source-fixtures-created", Count = manifest.Items.Count, manifest.SourceStoreId, manifest.SourceFolderId });
            // All preparation COM references have been released by the helper.
            // Diagnostic collection records body differences; it does NOT approve
            // those differences or allow source removal.
            journal.Write(new { State = "local-pilot-started", Transport = "unicode-msg-direct-import", Count = FixtureCount });
            int exitCode = PilotCommand.Execute(
                [manifest.SourceStoreId!, manifest.SourceFolderId!, targetPst, pilotRun, FixtureCount.ToString(System.Globalization.CultureInfo.InvariantCulture)],
                collectBodyDifferences: true, previousRun: null,
                useMsgTransport: true, allowInitialDiagnosticCollection: true);
            journal.Write(new { State = "local-pilot-returned", ExitCode = exitCode, RemovalAllowed = false });
            WriteJsonNew(Path.Combine(directory, "outcome.json"), new
            {
                FinishedUtc = DateTimeOffset.UtcNow, EngineExitCode = exitCode,
                SourceIsSyntheticLocalPst = true, DestinationIsNewLocalPst = true,
                RemovalAllowed = false, ProductionApproval = false,
                Scope = "Teste local artificial; consulte pilot-run/report.json e o diário. Divergência de corpo continua pendente de revisão."
            });
            Console.WriteLine($"Teste artificial local concluído com código {exitCode}. " +
                $"Nenhuma mensagem foi enviada. PSTs locais permanecem anexados. Evidências: {directory}");
            return exitCode; // In particular, preserve engine code 3 for pending body evidence.
        }
        catch (Exception ex)
        {
            manifest.Status = "review-required";
            Console.Error.WriteLine("Teste artificial interrompido: " + ex.GetType().Name + ": " + ex.Message);
            if (journal is not null)
            {
                try { journal.Write(new { State = "fixture-check-failed", ErrorType = ex.GetType().Name, HResult = $"0x{ex.HResult:X8}" }); }
                catch (Exception) { Console.Error.WriteLine("Não foi possível completar o diário local."); }
            }
            if (directoryCreated && directory is not null)
            {
                try
                {
                    // Preserve any published manifest. Partial fixture sets are
                    // evidence, never reused or deleted automatically.
                    if (!File.Exists(Path.Combine(directory, "manifest.json")))
                        WriteJsonNew(Path.Combine(directory, "manifest.json"), manifest);
                }
                catch (Exception) { Console.Error.WriteLine("Não foi possível completar o manifesto local."); }
            }
            return 1;
        }
        finally { journal?.Dispose(); }
    }

    private static void CreateFixtureMessages(string sourcePst, string inputs, DurableJournal journal, FixtureManifest manifest)
    {
        object? app = null, session = null, store = null, root = null, folders = null, folder = null, items = null;
        try
        {
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            if (File.Exists(sourcePst)) throw new IOException("PST de origem já existe; criação bloqueada.");
            journal.Write(new { State = "local-source-pst-create-started", Path = sourcePst });
            ((dynamic)session).AddStoreEx(sourcePst, 2);
            store = FindNewPst(session, sourcePst) ?? throw new InvalidOperationException("PST local criado não foi localizado.");
            string storeId = (string)((dynamic)store).StoreID;
            manifest.SourceStoreId = storeId;
            root = ((dynamic)store).GetRootFolder();
            if ((string)((dynamic)root).StoreID != storeId)
                throw new InvalidOperationException("Raiz não pertence ao PST artificial.");
            folders = ((dynamic)root).Folders;
            folder = ((dynamic)folders).Add("Mensagens artificiais", 6);
            string folderId = (string)((dynamic)folder).EntryID;
            if ((string)((dynamic)folder).StoreID != storeId)
                throw new InvalidOperationException("Pasta não pertence ao PST artificial.");
            manifest.SourceFolderId = folderId;
            items = ((dynamic)folder).Items;
            if ((int)((dynamic)items).Count != 0) throw new InvalidOperationException("A pasta artificial nova não está vazia.");

            for (int index = 1; index <= FixtureCount; index++)
            {
                object? item = null, staged = null, imported = null;
                try
                {
                    journal.Write(new { State = "fixture-create-started", Index = index });
                    item = ((dynamic)items).Add(0); // Guard the returned item before any configuration or Save.
                    GuardUnsavedSyntheticItem(item);
                    dynamic mail = item;
                    if ((int)mail.Class != 43 || (string)mail.MessageClass != "IPM.Note")
                        throw new InvalidOperationException("O provedor não criou uma mensagem IPM.Note local.");
                    mail.To = FixtureAddress; // Intentionally invalid TLD; never Resolve/Send.
                    mail.Subject = $"Teste local Unicode {index:D2} — ação e café";
                    string kind;
                    switch (index)
                    {
                        case 1:
                            kind = "plain-unicode-no-attachments";
                            mail.BodyFormat = 1;
                            mail.Body = "Mensagem artificial em texto simples.\r\nAção, café, coração, € e 漢字.\r\n";
                            break;
                        case 2:
                            kind = "plain-unicode-text-and-64kb-binary";
                            mail.BodyFormat = 1;
                            mail.Body = "Mensagem artificial com dois anexos locais e nomes com acentuação.\r\n";
                            AddAttachment(item, Path.Combine(inputs, "anexo-ação.txt"));
                            AddAttachment(item, Path.Combine(inputs, "dados-binários-64KB.bin"));
                            break;
                        case 3:
                            kind = "static-html-inline-png";
                            mail.BodyFormat = 2;
                            mail.HTMLBody = "<html><head><meta charset=\"utf-8\"></head><body>" +
                                "<p>Teste local: ação, café e coração.</p>" +
                                "<img src=\"cid:fixture-inline-image\" alt=\"Imagem artificial local\"></body></html>";
                            AddAttachment(item, Path.Combine(inputs, "imagem-ícone.png"), inline: true);
                            break;
                        case 4:
                            kind = "rtf-unicode-text-attachment";
                            mail.Body = "Teste local RTF: ação, café e informação.\r\n";
                            mail.BodyFormat = 3;
                            AddAttachment(item, Path.Combine(inputs, "relatório-RTF.txt"));
                            break;
                        default:
                            kind = "plain-two-megabyte-binary";
                            mail.BodyFormat = 1;
                            mail.Body = "Mensagem artificial com anexo binário determinístico de 2 MiB.\r\n";
                            AddAttachment(item, Path.Combine(inputs, "amostra-binária-2MB.bin"));
                            break;
                    }
                    GuardUnsavedSyntheticItem(item);
                    string preparedMsg = Path.Combine(inputs, $"prepared-fixture-{index:D3}.msg");
                    journal.Write(new { State = "fixture-msg-save-started", Index = index, Kind = kind });
                    mail.SaveAs(preparedMsg, 9); // File only: never Save this virtual Outbox object.
                    GuardUnsavedSyntheticItem(item);
                    SyntheticReceivedMsg.Prepare(preparedMsg); // Synthetic file only; no message is sent.
                    journal.Write(new { State = "fixture-local-import-started", Index = index, Kind = kind });
                    staged = ((dynamic)session).OpenSharedItem(preparedMsg);
                    imported = ((dynamic)staged).Move(folder);
                    GuardLocalParent(imported, storeId, folderId);
                    if (!(bool)((dynamic)imported).Sent) throw new InvalidDataException("Fixture não reproduz os sinalizadores de uma mensagem já recebida.");
                    string entryId = (string)((dynamic)imported).EntryID;
                    if (string.IsNullOrWhiteSpace(entryId)) throw new InvalidOperationException("Mensagem local salva sem EntryID.");
                    manifest.Items.Add(new(index, kind, entryId));
                    journal.Write(new { State = "fixture-saved", Index = index, Kind = kind, EntryId = entryId });
                }
                finally { Release(imported); Release(staged); Release(item); }
            }
            if ((int)((dynamic)items).Count != FixtureCount)
                throw new InvalidOperationException("A contagem da pasta artificial mudou durante a preparação.");
        }
        finally { Release(items); Release(folder); Release(folders); Release(root); Release(store); Release(session); Release(app); }
    }

    private static void GuardLocalParent(object item, string storeId, string folderId)
    {
        object? parent = null;
        try
        {
            parent = ((dynamic)item).Parent;
            if ((string)((dynamic)parent).StoreID != storeId || (string)((dynamic)parent).EntryID != folderId)
                throw new InvalidOperationException("Gravação bloqueada: item fora da pasta local artificial.");
        }
        finally { Release(parent); }
    }

    private static void GuardUnsavedSyntheticItem(object item)
    {
        if ((int)((dynamic)item).Class != 43 || (string)((dynamic)item).MessageClass != "IPM.Note" ||
            !string.IsNullOrEmpty((string?)((dynamic)item).EntryID))
            throw new InvalidOperationException("Objeto artificial não está exclusivamente em memória; nenhuma gravação será feita.");
    }

    private static void AddAttachment(object item, string path, bool inline = false)
    {
        object? attachments = null, attachment = null, accessor = null;
        try
        {
            attachments = ((dynamic)item).Attachments;
            attachment = ((dynamic)attachments).Add(path, 1, Type.Missing, Path.GetFileName(path));
            if (inline)
            {
                accessor = ((dynamic)attachment).PropertyAccessor;
                ((dynamic)accessor).SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "fixture-inline-image");
                ((dynamic)accessor).SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", true);
            }
        }
        finally { Release(accessor); Release(attachment); Release(attachments); }
    }

    private static object? FindNewPst(object session, string expectedPath)
    {
        object? stores = null;
        try
        {
            stores = ((dynamic)session).Stores;
            for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
            {
                object? candidate = null;
                try
                {
                    candidate = ((dynamic)stores)[i];
                    string path = (string)((dynamic)candidate).FilePath;
                    if (!string.IsNullOrWhiteSpace(path) && string.Equals(Path.GetFullPath(path), expectedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        if ((int)((dynamic)candidate).ExchangeStoreType != 3)
                            throw new InvalidOperationException("Origem artificial não é um armazenamento local não-Exchange.");
                        object found = candidate; candidate = null; return found;
                    }
                }
                catch (Exception ex) when (ex is COMException or ArgumentException) { }
                finally { Release(candidate); }
            }
            return null;
        }
        finally { Release(stores); }
    }

    private static byte[] PatternBytes(int length)
    {
        var bytes = new byte[length];
        for (int i = 0; i < bytes.Length; i++) bytes[i] = (byte)((i * 31 + 7) & 255);
        return bytes;
    }
    private static void WriteInput(string directory, string name, byte[] bytes, FixtureManifest manifest)
    {
        using (var stream = new FileStream(Path.Combine(directory, name), FileMode.CreateNew, FileAccess.Write, FileShare.None))
        { stream.Write(bytes); stream.Flush(true); }
        manifest.Inputs.Add(new(name, bytes.Length, Convert.ToHexString(SHA256.HashData(bytes))));
    }

    private static byte[] CreatePng()
    {
        using var png = new MemoryStream();
        png.Write(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 });
        byte[] ihdr = new byte[13];
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(0, 4), 1);
        BinaryPrimitives.WriteUInt32BigEndian(ihdr.AsSpan(4, 4), 1);
        ihdr[8] = 8; ihdr[9] = 6; // One transparent RGBA pixel, no external resources.
        WritePngChunk(png, "IHDR", ihdr);
        using var compressed = new MemoryStream();
        using (var zlib = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            zlib.Write(new byte[] { 0, 0, 0, 0, 0 }); // filter byte plus RGBA
        WritePngChunk(png, "IDAT", compressed.ToArray());
        WritePngChunk(png, "IEND", []);
        return png.ToArray();
    }
    private static void WritePngChunk(Stream output, string type, byte[] data)
    {
        byte[] name = Encoding.ASCII.GetBytes(type);
        Span<byte> number = stackalloc byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(number, (uint)data.Length);
        output.Write(number); output.Write(name); output.Write(data);
        uint crc = 0xFFFFFFFF;
        foreach (byte value in name.Concat(data))
        {
            crc ^= value;
            for (int bit = 0; bit < 8; bit++) crc = (crc >> 1) ^ ((crc & 1) != 0 ? 0xEDB88320u : 0u);
        }
        BinaryPrimitives.WriteUInt32BigEndian(number, crc ^ 0xFFFFFFFF);
        output.Write(number);
    }

    private static void WriteJsonNew<T>(string output, T value)
    {
        string temp = output + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            { JsonSerializer.Serialize(stream, value, new JsonSerializerOptions { WriteIndented = true }); stream.Flush(true); }
            File.Move(temp, output, overwrite: false);
        }
        finally
        {
            try { File.Delete(temp); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }
    private static void CreateDirectoryExclusive(string path)
    {
        if (!CreateDirectoryW(path, IntPtr.Zero))
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Não foi possível criar um diretório novo exclusivo.");
    }
    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateDirectoryW(string path, IntPtr securityAttributes);
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);

    private sealed record InputArtifact(string Name, int Bytes, string Sha256);
    private sealed record FixtureItem(int Index, string Kind, string EntryId);
    private sealed class FixtureManifest
    {
        public int SchemaVersion { get; set; } = 1;
        public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
        public string Status { get; set; } = "preparing";
        public bool SyntheticOnly { get; } = true;
        public bool SendsEmail { get; } = false;
        public bool RemovalAllowed { get; } = false;
        public string Recipient { get; } = FixtureAddress;
        public string? SourcePstPath { get; set; }
        public string? TargetPstPath { get; set; }
        public string? PilotRunDirectory { get; set; }
        public string? SourceStoreId { get; set; }
        public string? SourceFolderId { get; set; }
        public List<InputArtifact> Inputs { get; set; } = new();
        public List<FixtureItem> Items { get; set; } = new();
    }
}
