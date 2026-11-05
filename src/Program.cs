using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        if (args.Length > 0 && args[0] == "native-batch-probe") return NativeBatchProbe.Execute(args.Skip(1).ToArray());
        if (args.Length > 0 && args[0] == "real-native-batch-probe") return RealNativeBatchProbe.Execute(args.Skip(1).ToArray());
        if (args.Length > 0 && args[0].StartsWith("archive-", StringComparison.Ordinal)) return ArchiveCommand.Execute(args);
        if (args.Length > 0 && args[0] == "pilot-copy") return PilotCommand.Execute(args.Skip(1).ToArray());
        if (args.Length > 0 && args[0] == "pilot-sample") return PilotCommand.Sample(args.Skip(1).ToArray());
        if (args.Length > 0 && args[0] == "pilot-msg-sample") return PilotCommand.Sample(args.Skip(1).ToArray(), true);
        if (args.Length == 3 && args[0] == "msg-diff") return MsgDifferenceAnalysis.Execute(args.Skip(1).ToArray());
        if (args.Length == 1 && args[0] == "pilot-self-test") return PilotCommand.SelfTest();
        if (args.Length == 1 && args[0] == "self-test") return SelfTest();
        if (args.Length == 2 && args[0] == "pilot-review") return PilotCommand.Review(args[1]);
        if (args.Length == 2 && args[0] == "pilot-reconcile-junk") return JunkReconciliation.Execute(args[1]);
        if (args.Length == 2 && args[0] == "pilot-audit") return PilotAudit.Execute(args[1]);
        if (args.Length == 2 && args[0] == "pst-reopen-check") return PstReopenCheck.Execute(args[1]);
        if (args.Length == 4 && args[0] == "folder-coverage") return FolderCoverage.Execute(args.Skip(1).ToArray());
        if (args.Length == 2 && args[0] == "local-fixture-check") return LocalFixtureCheck.Execute(args.Skip(1).ToArray());
        if (args.Length == 2 && args[0] == "fixture-parent-probe") return LocalFixtureCheck.Probe(args[1]);
        if (args.Length == 0 || args[0] is "help" or "--help")
        {
            Console.WriteLine("Arquivamento Outlook — inventário e cópia piloto\n" +
                "  list-stores\n  inventory <store-id> <arquivo.json>\n  self-test\n" +
                "  pilot-copy <store-id> <folder-entry-id> <destino.pst> <diretorio-execucao> [limite-1-a-25]\n" +
                "  pilot-sample <store-id> <folder-entry-id> <novo.pst> <nova-execucao> <execucao-anterior> [limite-1-a-25]\n" +
                "  pilot-msg-sample <store-id> <folder-entry-id> <novo.pst> <nova-execucao> <execucao-anterior> [limite-1-a-25]\n" +
                "  msg-diff <origem.msg> <destino.msg> (diagnóstico offline)\n" +
                "  pilot-self-test\n" +
                "  archive-plan <config.json>\n  archive-run|archive-verify <raiz> <execucao>\n" +
                "  archive-remove <raiz> <execucao> deleted-items|mailbox\n  archive-status <raiz> [execucao]\n  archive-self-test\n" +
                "  pilot-review <diretorio-execucao>\n" +
                "  pilot-audit <diretorio-execucao> (auditoria offline)\n" +
                "  pst-reopen-check <diretorio-execucao> (desanexa/reabre somente o PST do piloto)\n" +
                "  folder-coverage <store-id> <folder-id> <novo.json> (somente leitura)\n" +
                "  local-fixture-check <novo-diretorio-local> (mensagens artificiais em PSTs novos)\n" +
                "  pilot-reconcile-junk <diretorio-execucao> (diagnóstico: move somente cópia dentro do PST)\n" +
                "Abra o Outlook clássico com o perfil desejado antes de executar.\n" +
                "Inventário somente leitura; o piloto cria PST e copia um lote limitado, sem remover originais.");
            return 0;
        }
        if (!(args.Length == 1 && args[0] == "list-stores") &&
            !(args.Length == 3 && args[0] == "inventory"))
        {
            Console.Error.WriteLine("Argumentos inválidos. Execute --help.");
            return 2;
        }
        object? app = null, session = null, stores = null;
        try
        {
            // Attach only: never launch another profile or close the user's Outlook.
            var clsid = new Guid("0006F03A-0000-0000-C000-000000000046");
            Marshal.ThrowExceptionForHR(GetActiveObject(ref clsid, IntPtr.Zero, out app));
            session = ((dynamic)app!).Session;
            stores = ((dynamic)session).Stores;
            if (args[0] == "list-stores")
            {
                var result = new StoreListing();
                for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
                {
                    object? store = null;
                    try { store = ((dynamic)stores)[i]; result.Stores.Add(Describe(store)); }
                    catch (Exception ex) { result.Errors.Add(new($"store-index:{i}", "describe-store", ex.Message)); }
                    finally { Release(store); }
                }
                Console.WriteLine(JsonSerializer.Serialize(result, JsonOptions));
                return result.Errors.Count == 0 ? 0 : 3;
            }
            var output = Path.GetFullPath(args[2]);
            if (File.Exists(output)) throw new IOException("O relatório já existe. Escolha outro nome.");
            if (!Directory.Exists(Path.GetDirectoryName(output))) throw new IOException("A pasta do relatório não existe.");
            object? selected = null, root = null;
            try
            {
                for (int i = 1; i <= (int)((dynamic)stores).Count; i++)
                {
                    object? candidate = null;
                    try
                    {
                        candidate = ((dynamic)stores)[i];
                        if (string.Equals((string)((dynamic)candidate).StoreID, args[1], StringComparison.Ordinal))
                        { selected = candidate; candidate = null; break; }
                    }
                    finally { Release(candidate); }
                }
                if (selected is null) throw new ArgumentException("Caixa não encontrada neste perfil.");
                var report = new InventoryReport { Store = Describe(selected) };
                root = ((dynamic)selected).GetRootFolder();
                Visit(root, "", report);
                report.FinishedUtc = DateTimeOffset.UtcNow;
                WriteReportAtomically(output, report);
                Console.WriteLine($"Inventário salvo em {output}. Pastas: {report.Folders.Count}; falhas: {report.Errors.Count}.");
                Console.WriteLine("O relatório não comprova completude no servidor nem autoriza remoção de itens.");
                return report.Errors.Count == 0 ? 0 : 3;
            }
            finally { Release(root); Release(selected); }
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800401E3))
        {
            Console.Error.WriteLine("O Outlook clássico não está acessível pela tabela de objetos COM. " +
                "Abra o clássico, aguarde o perfil terminar de carregar e alterne para outra janela antes de tentar novamente. " +
                "Confirme também o mesmo usuário e nível de permissão. O programa não iniciará outro perfil automaticamente.");
            return 1;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Falha: {ex.Message}");
            return 1;
        }
        finally { Release(stores); Release(session); Release(app); }
    }

    private static void WriteReportAtomically(string output, InventoryReport report)
    {
        // Same directory ensures that publishing the completed file is a rename,
        // not a cross-volume copy. Never replace an existing report.
        var temporary = Path.Combine(Path.GetDirectoryName(output)!, $".{Path.GetFileName(output)}.{Guid.NewGuid():N}.tmp");
        try
        {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                JsonSerializer.Serialize(stream, report, JsonOptions);
                stream.Flush(flushToDisk: true);
            }
            File.Move(temporary, output, overwrite: false);
        }
        finally
        {
            try { File.Delete(temporary); }
            catch (IOException) { /* Keep the original reporting failure. */ }
            catch (UnauthorizedAccessException) { /* A leftover temporary file is not a completed report. */ }
        }
    }

    private static void Visit(object folder, string parent, InventoryReport report)
    {
        object? children = null, items = null;
        string path = parent;
        try
        {
            dynamic f = folder;
            path = parent + "/" + (string)f.Name;
            // Virtual search folders overlap real folders; never count them twice.
            const string folderType = "http://schemas.microsoft.com/mapi/proptag/0x36010003";
            object? accessor = null;
            try
            {
                accessor = f.PropertyAccessor;
                if ((int)((dynamic)accessor).GetProperty(folderType) == 2)
                { report.Folders.Add(new(path, (string)f.EntryID, null, "search-folder-excluded")); return; }
            }
            finally { Release(accessor); }
            int? count = null;
            try { items = f.Items; count = (int)((dynamic)items).Count; }
            catch (Exception ex) { report.Errors.Add(new(path, "count", ex.Message)); }
            report.Folders.Add(new(path, (string)f.EntryID, count, count is null ? "error" : "observed"));
            children = f.Folders;
            for (int i = 1; i <= (int)((dynamic)children).Count; i++)
            {
                object? child = null;
                try { child = ((dynamic)children)[i]; Visit(child, path, report); }
                catch (Exception ex) { report.Errors.Add(new(path, "child", ex.Message)); }
                finally { Release(child); }
            }
        }
        catch (Exception ex) { report.Errors.Add(new(path, "folder", ex.Message)); }
        finally { Release(items); Release(children); }
    }

    private static StoreInfo Describe(dynamic store) => new((string)store.DisplayName,
        (string)store.StoreID, (bool)store.IsDataFileStore, (bool)store.IsCachedExchange,
        (string)store.FilePath, (int)store.ExchangeStoreType);
    private static void Release(object? value)
    { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? instance);

    private static int SelfTest()
    {
        var report = new InventoryReport { Store = new("Teste com acentuação", "fake", false, true) };
        report.Folders.Add(new("/Caixa/Enviados", "folder1", 12, "observed"));
        report.Folders.Add(new("/Caixa/Pesquisa", "folder2", null, "search-folder-excluded"));
        report.Errors.Add(new("/Caixa/Pendente", "count", "Falha simulada"));
        var json = JsonSerializer.Serialize(report, JsonOptions);
        var restored = JsonSerializer.Deserialize<InventoryReport>(json)!;
        if (restored.CompletenessVerified || restored.Folders.Count != 2 ||
            restored.Errors.Count != 1 || restored.Folders[1].ItemCount is not null ||
            restored.Store?.Name != report.Store.Name) throw new Exception("Falha no contrato do relatório.");
        Console.WriteLine("Teste do formato do relatório aprovado. Integração COM não testada.");
        return 0;
    }
}

// IsDataFileStore also includes OST; neither it nor FilePath proves that a path is local.
internal record StoreInfo(string Name, string StoreId, bool IsDataFileStore, bool IsCachedExchange,
    string? FilePath = null, int? ExchangeStoreType = null);
internal sealed class StoreListing
{
    public int SchemaVersion { get; set; } = 1;
    public List<StoreInfo> Stores { get; set; } = new();
    public List<InventoryError> Errors { get; set; } = new();
}
internal record FolderInfo(string Path, string EntryId, int? ItemCount, string Status);
internal record InventoryError(string Folder, string Stage, string Message);
internal sealed class InventoryReport
{
    public int SchemaVersion { get; set; } = 1;
    public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset? FinishedUtc { get; set; }
    public StoreInfo? Store { get; set; }
    public bool CompletenessVerified { get; set; } = false;
    public string Scope { get; set; } = "Itens normais acessíveis pelo Outlook; sem garantia de completude no servidor.";
    public List<FolderInfo> Folders { get; set; } = new();
    public List<InventoryError> Errors { get; set; } = new();
}
