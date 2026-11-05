using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using OpenMcdf;

// Factory-only adjustment of a NEW ARTIFICIAL MSG file, before local PST import.
// This clears MSGFLAG_UNSENT (0x8) only. It does not send mail, set a submission
// flag, contact Outlook, edit a PST, or modify delivery/submit timestamps.
internal static class SyntheticReceivedMsg
{
    private const uint MessageFlagsTag = 0x0E070003;
    private const uint UnsentMask = 0x00000008;
    private const string PropertiesName = "__properties_version1.0";
    private const int MaximumPropertiesBytes = 1024 * 1024;

    internal static void Prepare(string msgPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(msgPath);
        if (!Path.IsPathFullyQualified(msgPath) || !File.Exists(msgPath) ||
            !string.Equals(Path.GetExtension(msgPath), ".msg", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("A preparação exige arquivo MSG artificial local existente com caminho absoluto.");
        if (msgPath.StartsWith(@"\\", StringComparison.Ordinal) ||
            new DriveInfo(Path.GetPathRoot(msgPath)!).DriveType != DriveType.Fixed)
            throw new ArgumentException("O arquivo artificial deve estar em disco local fixo.");
        if (File.GetAttributes(msgPath).HasFlag(FileAttributes.ReparsePoint))
            throw new ArgumentException("Arquivo MSG redirecionado não é permitido.");
        for (var directory = new DirectoryInfo(Path.GetDirectoryName(msgPath)!); directory is not null; directory = directory.Parent)
            if (directory.Attributes.HasFlag(FileAttributes.ReparsePoint))
                throw new ArgumentException("Caminho do MSG artificial contém redirecionamento.");

        byte[] before;
        int flagOffset;
        using (var read = RootStorage.OpenRead(msgPath))
            (before, flagOffset) = ReadValidatedProperties(read);
        uint oldFlags = BinaryPrimitives.ReadUInt32LittleEndian(before.AsSpan(flagOffset, 4));
        uint newFlags = oldFlags & ~UnsentMask;
        if (newFlags == oldFlags) return; // Idempotent; do not open writable unnecessarily.
        byte[] expected = (byte[])before.Clone();
        BinaryPrimitives.WriteUInt32LittleEndian(expected.AsSpan(flagOffset, 4), newFlags);

        using (var write = RootStorage.Open(msgPath, FileMode.Open, FileAccess.ReadWrite, StorageModeFlags.None))
        {
            var (current, currentOffset) = ReadValidatedProperties(write);
            if (currentOffset != flagOffset || !current.AsSpan().SequenceEqual(before))
                throw new InvalidDataException("MSG artificial mudou antes da preparação; gravação bloqueada.");
            using (var stream = write.OpenStream(PropertiesName))
            {
                stream.Position = flagOffset;
                stream.Write(expected.AsSpan(flagOffset, 4)); // Only the uint32 value of PR_MESSAGE_FLAGS.
                stream.Flush();
            }
            write.Flush();
        }
        // Reopen independently after flushing/disposal. All other property bytes,
        // including fixed-length FILETIMEs and entry flags, must remain unchanged.
        using var verified = RootStorage.OpenRead(msgPath);
        var (after, verifiedOffset) = ReadValidatedProperties(verified);
        if (verifiedOffset != flagOffset || !after.AsSpan().SequenceEqual(expected))
            throw new InvalidDataException("Preparação do MSG artificial não preservou as demais propriedades.");
    }

    private static (byte[] Properties, int FlagOffset) ReadValidatedProperties(RootStorage root)
    {
        using var properties = root.OpenStream(PropertiesName);
        if (properties.Length < 32 || properties.Length > MaximumPropertiesBytes || (properties.Length - 32) % 16 != 0)
            throw new InvalidDataException("Cabeçalho/alinhamento da tabela de propriedades MSG inválido.");
        byte[] bytes = new byte[checked((int)properties.Length)];
        properties.ReadExactly(bytes);
        int flagOffset = -1;
        int classCount = 0;
        uint classTag = 0;
        uint declaredClassSize = 0;
        for (int offset = 32; offset < bytes.Length; offset += 16)
        {
            uint tag = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
            if (tag == MessageFlagsTag)
            {
                if (flagOffset >= 0) throw new InvalidDataException("PR_MESSAGE_FLAGS duplicado.");
                flagOffset = offset + 8;
            }
            if (tag is 0x001A001F or 0x001A001E)
            {
                classCount++;
                classTag = tag;
                declaredClassSize = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 8, 4));
            }
        }
        if (flagOffset < 0) throw new InvalidDataException("PR_MESSAGE_FLAGS ausente.");
        if (classCount != 1) throw new InvalidDataException("MessageClass deve existir exatamente uma vez.");
        using var classStream = root.OpenStream("__substg1.0_" + classTag.ToString("X8"));
        int terminatorBytes = classTag == 0x001A001F ? 2 : 1;
        if (classStream.Length is < 1 or > 1024 || declaredClassSize != classStream.Length + terminatorBytes)
            throw new InvalidDataException("Comprimento de MessageClass inválido.");
        byte[] encodedClass = new byte[checked((int)classStream.Length)];
        classStream.ReadExactly(encodedClass);
        string messageClass = (classTag == 0x001A001F ? new UnicodeEncoding(false, false, true) : Encoding.ASCII)
            .GetString(encodedClass).TrimEnd('\0');
        if (messageClass != "IPM.Note") throw new InvalidDataException("O arquivo não representa uma mensagem IPM.Note artificial.");
        return (bytes, flagOffset);
    }

    internal static int SelfTest()
    {
        string directory = Path.Combine(Path.GetTempPath(), "outlook-synthetic-flags-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string[] names = ["flags8.msg", "flags9.msg", "duplicate.msg", "missing.msg", "alignment.msg", "wrong-class.msg"];
        try
        {
            foreach (uint flags in new uint[] { 8, 9 })
            {
                string path = Path.Combine(directory, $"flags{flags}.msg");
                WriteFixture(path, flags);
                byte[] before, beforeBody;
                int offset;
                using (var root = RootStorage.OpenRead(path))
                {
                    (before, offset) = ReadValidatedProperties(root);
                    using var body = root.OpenStream("__substg1.0_1000001F");
                    beforeBody = new byte[checked((int)body.Length)]; body.ReadExactly(beforeBody);
                }
                Prepare(path);
                using (var root = RootStorage.OpenRead(path))
                {
                    var (after, afterOffset) = ReadValidatedProperties(root);
                    BinaryPrimitives.WriteUInt32LittleEndian(before.AsSpan(offset, 4), flags & ~UnsentMask);
                    if (afterOffset != offset || !before.AsSpan().SequenceEqual(after))
                        throw new Exception("Preparação alterou datas/outras propriedades da fixture.");
                    using var body = root.OpenStream("__substg1.0_1000001F");
                    byte[] afterBody = new byte[checked((int)body.Length)]; body.ReadExactly(afterBody);
                    if (!beforeBody.AsSpan().SequenceEqual(afterBody)) throw new Exception("Corpo da fixture foi alterado.");
                }
                string hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
                Prepare(path);
                if (hash != Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))))
                    throw new Exception("Preparação idempotente regravou a fixture.");
            }
            WriteFixture(Path.Combine(directory, "duplicate.msg"), 8, duplicate: true);
            WriteFixture(Path.Combine(directory, "missing.msg"), 8, missing: true);
            WriteFixture(Path.Combine(directory, "alignment.msg"), 8, misaligned: true);
            WriteFixture(Path.Combine(directory, "wrong-class.msg"), 8, wrongClass: true);
            foreach (string name in names.Skip(2))
            {
                string path = Path.Combine(directory, name);
                string before = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));
                bool rejected = false;
                try { Prepare(path); } catch (InvalidDataException) { rejected = true; }
                if (!rejected || before != Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))))
                    throw new Exception("Fixture inválida não foi recusada antes da gravação.");
            }
            Console.WriteLine("MSG artificial: flags 8→0 e 9→1; datas/corpo preservados; idempotência e rejeição de tabelas inválidas aprovadas. Sem Outlook.");
            return 0;
        }
        finally
        {
            foreach (string name in names) File.Delete(Path.Combine(directory, name));
            Directory.Delete(directory);
        }
    }

    private static void WriteFixture(string path, uint flags, bool duplicate = false, bool missing = false,
        bool misaligned = false, bool wrongClass = false)
    {
        using var root = RootStorage.Create(path);
        byte[] classBytes = Encoding.Unicode.GetBytes(wrongClass ? "IPM.Task" : "IPM.Note");
        byte[] bodyBytes = Encoding.Unicode.GetBytes("Conteúdo artificial: ação.");
        using (var value = root.CreateStream("__substg1.0_001A001F")) value.Write(classBytes);
        using (var value = root.CreateStream("__substg1.0_1000001F")) value.Write(bodyBytes);
        using var table = root.CreateStream(PropertiesName);
        table.Write(new byte[32]);
        void Entry(uint tag, ulong value)
        {
            Span<byte> entry = stackalloc byte[16];
            entry.Clear();
            BinaryPrimitives.WriteUInt32LittleEndian(entry[..4], tag);
            BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(4, 4), 6);
            BinaryPrimitives.WriteUInt64LittleEndian(entry.Slice(8, 8), value);
            table.Write(entry);
        }
        Entry(0x001A001F, (ulong)classBytes.Length + 2);
        Entry(0x1000001F, (ulong)bodyBytes.Length + 2);
        Entry(0x0E060040, (ulong)new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc).ToFileTimeUtc());
        Entry(0x00390040, (ulong)new DateTime(2020, 1, 2, 2, 3, 4, DateTimeKind.Utc).ToFileTimeUtc());
        Entry(0x00360003, 2);
        if (!missing) Entry(MessageFlagsTag, flags);
        if (duplicate) Entry(MessageFlagsTag, flags);
        if (misaligned) table.WriteByte(0);
    }
}
