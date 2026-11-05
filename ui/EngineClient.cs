using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace OutlookArchive.UI;

internal sealed record EngineResult(int ExitCode, string Output, string Error, bool WasInterrupted = false);
internal sealed class EngineInterruptedException : Exception;
internal sealed class EngineClient
{
    private sealed class RunState(Process process) { internal Process Process = process; internal bool Interrupted; }
    private readonly object gate = new();
    private RunState? running;
    internal bool IsRunning { get { lock (gate) return running?.Process is { HasExited: false }; } }
    internal string EnginePath { get; }
    internal string? LogDirectory { get; set; }
    internal EngineClient(string? enginePath = null)
    {
        EnginePath = enginePath ?? LocateEngine();
    }
    private static string LocateEngine()
    {
        string beside = Path.Combine(AppContext.BaseDirectory, "engine", "OutlookArchive.dll");
        if (File.Exists(beside)) return beside;
        DirectoryInfo? directory = new(AppContext.BaseDirectory);
        while (directory is not null)
        {
            foreach (string configuration in new[] { "Release", "Debug" })
            {
                string candidate = Path.Combine(directory.FullName, "src", "bin", configuration, "net8.0-windows", "OutlookArchive.dll");
                if (File.Exists(candidate)) return candidate;
            }
            directory = directory.Parent;
        }
        return beside;
    }
    internal async Task<EngineResult> RunAsync(IEnumerable<string> arguments, Action<string>? output = null)
    {
        arguments = arguments.ToArray();
        if (arguments.FirstOrDefault() is "archive-run" or "archive-verify" or "archive-workflow-run" or "archive-remove" or "archive-reconcile" or "archive-plan")
        {
            string directory = Path.GetDirectoryName(EnginePath)!;
            var missing = new[] { "native-copy.exe", "native-identity.exe", "native-finalize.exe" }.Where(x => !File.Exists(Path.Combine(directory, x))).ToArray();
            if (missing.Length > 0) throw new InvalidOperationException("Instalação incompleta: faltam os componentes " + string.Join(", ", missing) + ". Abra o programa pelo iniciador para recompilar com os componentes necessários. Nenhuma operação foi iniciada.");
        }
        if (IsRunning) throw new InvalidOperationException("Já existe uma operação em andamento nesta janela.");
        if (!File.Exists(EnginePath)) throw new FileNotFoundException("O motor de arquivamento não foi encontrado. Abra pelo iniciador da ferramenta.", EnginePath);
        var info = new ProcessStartInfo("dotnet")
        {
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8,
            WorkingDirectory = Path.GetDirectoryName(EnginePath)!
        };
        info.ArgumentList.Add(EnginePath);
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        using var process = new Process { StartInfo = info };
        var stdout = new StringBuilder(); var stderr = new StringBuilder();
        using var operationLog=LogDirectory is null ? null : CreateLog(LogDirectory,arguments.FirstOrDefault() ?? "operation");
        void Log(string line){if(operationLog is not null)lock(operationLog){operationLog.WriteLine(line);operationLog.Flush();}}
        process.OutputDataReceived += (_, e) => { if (e.Data is not null) { lock (stdout) stdout.AppendLine(e.Data); Log(e.Data);output?.Invoke(e.Data); } };
        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            Log(e.Data);
            lock (stderr)
            {
                if (stderr.Length > 90_000) stderr.Remove(0, stderr.Length - 60_000);
                stderr.AppendLine(e.Data);
            }
            output?.Invoke(e.Data);
        };
        try { if (!process.Start()) throw new IOException("Não foi possível iniciar o motor."); }
        catch (System.ComponentModel.Win32Exception ex) { throw new InvalidOperationException("Não foi possível iniciar o .NET. Abra pelo iniciador e confirme a instalação do .NET 8.", ex); }
        var state = new RunState(process);
        lock (gate) running = state;
        try
        {
            process.BeginOutputReadLine(); process.BeginErrorReadLine();
            await process.WaitForExitAsync();
            process.WaitForExit(); // Drain asynchronous output events before parsing JSON.
            lock (gate) return new(process.ExitCode, stdout.ToString(), stderr.ToString(), state.Interrupted);
        }
        finally { lock (gate) { if (ReferenceEquals(running, state)) running = null; } }
    }
    private static StreamWriter CreateLog(string directory,string command)
    {
        Directory.CreateDirectory(directory);
        return new StreamWriter(Path.Combine(directory,DateTime.Now.ToString("yyyyMMdd-HHmmss-fff")+"-"+command+"-"+Guid.NewGuid().ToString("N")[..8]+".log"),false,new UTF8Encoding(false));
    }
    internal void Interrupt()
    {
        lock (gate)
        {
            try
            {
                if (running is { } state && !state.Process.HasExited)
                { state.Process.Kill(entireProcessTree: false); state.Interrupted = true; }
            }
            catch (InvalidOperationException) { /* The process completed before Kill; preserve its actual result. */ }
            catch (System.ComponentModel.Win32Exception) { /* No successful Kill: do not classify a real failure as cancellation. */ }
        }
    }
}

internal sealed class SourceListing { public List<SourceStore> Stores { get; set; } = []; public List<string> Warnings { get; set; } = []; }
internal sealed class SourceStore
{
    public string Name { get; set; } = "";
    public string StoreId { get; set; } = "";
    public List<SourceFolder> Folders { get; set; } = [];
}
internal sealed class SourceFolder
{
    public string Name { get; set; } = "";
    public string FolderId { get; set; } = "";
    public string Path { get; set; } = "";
}
internal sealed record SourceSelection(string StoreId, string FolderId, string Label);
internal sealed class JobStatus
{
    public bool PlanningPending {get;set;}
    public int PlanningAnalyzed {get;set;}
    public int PlanningTotal {get;set;}
    public string Id { get; set; } = "";
    public DateTimeOffset CreatedUtc { get; set; }
    public string? SourceLabel { get; set; }
    public bool RetainedMsgBackupForRemoval { get; set; }
    public string? ExistingPstPath { get; set; }
    public string? DestinationPstPath { get; set; }
    public bool MoveValidatedToDeletedItems { get; set; }
    public WorkflowInfo? Workflow {get;set;}
    public bool PreserveFolderStructure { get; set; }
    public DateOnly? CutoffDateInclusive { get; set; }
    public string? CutoffTimeZoneId { get; set; }
    public Dictionary<string, int> States { get; set; } = [];
    public List<VolumeStatus> Volumes { get; set; } = [];
    public List<JobError> Errors { get; set; } = [];
    public override string ToString() => $"{CreatedUtc.ToLocalTime():dd/MM/yyyy HH:mm} · {SourceLabel??Id[..Math.Min(8,Id.Length)]} · {(PlanningPending?"planejamento interrompido":CutoffDateInclusive is { } cutoff?$"até {cutoff:dd/MM/yyyy}":"todos os e-mails")} · {Id[..Math.Min(8,Id.Length)]}";
}
internal sealed class JobError
{
    public string Key { get; set; } = "";
    public string Error { get; set; } = "";
}
internal sealed class NewItemsCheck
{
    public required int NewItemsCount { get; set; }
    public DateTimeOffset CheckedUtc { get; set; }
    public string? Error { get; set; }
}
internal sealed record NewItemsFollowUpResult(int NewItemsCount = 0, string? NewJobId = null, string? Warning = null, bool WasInterrupted = false);
internal static class NewItemsFollowUp
{
    internal static bool ShouldCheck(string command, EngineResult result) =>
        command is "archive-run" or "archive-verify" or "archive-workflow-run" && result.ExitCode is 0 or 3 && !result.WasInterrupted;

    internal static async Task<NewItemsFollowUpResult> ExecuteAsync(string directory, string previousJob,
        Func<string[], Task<EngineResult>> run, Func<int, Task<bool>> askToPrepare)
    {
        int count = 0;
        try
        {
            var response = await run(["archive-check-new", directory, previousJob]);
            if (response.WasInterrupted) return new(WasInterrupted: true);
            if (response.ExitCode != 0 && !response.Output.TrimStart().StartsWith("{", StringComparison.Ordinal))
                return new(Warning: "Não foi possível consultar novos e-mails. " + Failure(response));
            var check = JsonSerializer.Deserialize<NewItemsCheck>(response.Output, UiRules.Json) ?? throw new JsonException("Resposta vazia na consulta de novos e-mails.");
            if (!string.IsNullOrWhiteSpace(check.Error)) return new(Warning: "Consulta de novos e-mails: " + check.Error);
            if (response.ExitCode != 0) return new(Warning: "Não foi possível consultar novos e-mails. " + Failure(response));
            if (check.NewItemsCount < 0) throw new JsonException("Contagem inválida na consulta de novos e-mails.");
            count = check.NewItemsCount;
            if (count == 0 || !await askToPrepare(count)) return new(count);
            var planned = await run(["archive-plan-new", directory, previousJob]);
            if (planned.WasInterrupted) return new(count, WasInterrupted: true);
            if (planned.ExitCode != 0) return new(count, Warning: "O novo plano não foi concluído. " + Failure(planned));
            string? job = planned.Output.Split('\n').Select(x => x.Trim()).FirstOrDefault(x => x.StartsWith("EXECUCAO=", StringComparison.Ordinal))?[9..];
            if (!Guid.TryParseExact(job, "N", out _)) return new(count, Warning: "O motor não informou uma identidade válida para o novo plano. Consulte o histórico antes de preparar novamente.");
            return new(count, job);
        }
        catch (Exception ex)
        {
            // This optional follow-up must not turn an already completed copy
            // or verification into a reported failure.
            return new(count, Warning: "Não foi possível concluir a consulta ou preparação dos novos e-mails: " + ex.Message);
        }
    }
    private static string Failure(EngineResult result) => string.IsNullOrWhiteSpace(result.Error)
        ? $"O motor retornou o código {result.ExitCode}." : result.Error.Trim();
}
internal sealed class VolumeStatus
{
    public int Number { get; set; }
    public string PstPath { get; set; } = "";
    public long ObservedBytes { get; set; }
    public string State { get; set; } = "";
}
internal static class UiRules
{
    internal static string ErrorMessage(string technicalMessage)
    {
        if(technicalMessage.Contains("0x800706BA",StringComparison.OrdinalIgnoreCase) ||
           technicalMessage.Contains("0x800401E3",StringComparison.OrdinalIgnoreCase))
            return "Abra o Outlook clássico e aguarde o carregamento da sua caixa de e-mail. Depois, volte a este programa e tente novamente.\n\nO novo Outlook não é compatível com esta automação.";
        return technicalMessage;
    }
    internal static string[] VerificationArguments(string directory, string job, bool allowRetainedBackup) => allowRetainedBackup
        ? ["archive-verify", directory, job, "--pending-only", "--allow-retained-msg-backup"]
        : ["archive-verify", directory, job, "--pending-only"];
    internal static bool CanPrepareRemoval(IReadOnlyDictionary<string, int> states) =>
        states.Any(x => x.Value > 0 && x.Key is "Copied" or "Verified" or "InDeletedItems" or "DeleteStarted" or "PurgeStarted");
    internal const string PendingGuide = "A cópia rápida registra a transferência em lote; Copiado não significa Validado. A conferência detalhada pode demorar e deve ser feita antes de remover originais.\n\n1. Selecione a execução e abra a aba Pendências. O relatório reúne as contagens e os erros registrados.\n\n2. Use Reconciliar pendências para reconhecer cópias já existentes. Essa etapa não repete o envio de lotes.\n\n3. Use Verificar pendentes para conferir apenas as cópias ainda não aprovadas. Os itens já validados serão conferidos novamente antes da exclusão. Se o PST estiver ocupado, feche mensagens ou janelas que estejam exibindo esse arquivo e tente novamente.\n\n4. Use Copiar / retomar para continuar os itens ainda não copiados do plano inicial. Não prepare outro plano para corrigir a mesma execução.\n\nDivergências de conteúdo, cópias ambíguas e lotes incompletos exigem revisão; os originais permanecem protegidos. Não apague cópias nem redefina o histórico para forçar uma nova tentativa. Rascunhos e tipos não suportados continuam pendentes.";
    internal static bool IsActionableError(JobError error) => !error.Error.StartsWith("Cópia rápida concluída;", StringComparison.Ordinal);
    internal static string PendingReport(JobStatus job)
    {
        // Quoted cells preserve line breaks/separators. Prefix formula-leading
        // input so a diagnostic string cannot become a spreadsheet formula.
        string Cell(string value)
        {
            if (value.TrimStart().StartsWith('=') || value.TrimStart().StartsWith('+') || value.TrimStart().StartsWith('-') || value.TrimStart().StartsWith('@')) value = "'" + value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        var lines = new List<string> { "Execução;Categoria;Quantidade;Item;Motivo;Orientação" };
        foreach (var state in job.States.OrderBy(x => x.Key, StringComparer.Ordinal))
            lines.Add(string.Join(";", new[] { job.Id, StateLabel(state.Key), state.Value.ToString(), "", "", state.Key == "Copied" ? "Use Continuar arquivamento para validar as cópias antes de remover originais. A conferência detalhada pode demorar." : "" }.Select(Cell)));
        foreach (var error in job.Errors.Where(IsActionableError))
            lines.Add(string.Join(";", new[] { job.Id, "Pendência registrada", "", error.Key, error.Error, PendingAction(error.Error) }.Select(Cell)));
        return string.Join(Environment.NewLine, lines) + Environment.NewLine;
    }
    internal static string PendingAction(string error)
    {
        if (error.Contains("origem não foi localizada", StringComparison.OrdinalIgnoreCase))
            return "Confira se o original foi movido ou excluído. Preserve o MSG; a ferramenta não trocará a mensagem por outra.";
        if (error.Contains("Nenhuma cópia sem identificador", StringComparison.OrdinalIgnoreCase))
            return "Não há cópia correspondente identificada na pasta prevista. Preserve o MSG e revise o destino antes de autorizar uma recuperação.";
        if (error.Contains("Cópia rápida concluída", StringComparison.OrdinalIgnoreCase))
            return "A transferência terminou. Use Verificar pendentes antes de remover os originais.";
        if (error.Contains("volume-reopen", StringComparison.OrdinalIgnoreCase) || error.Contains("PST ocupado", StringComparison.OrdinalIgnoreCase))
            return "Feche janelas que exibem este PST e use Verificar pendentes.";
        if (error.Contains("relevantes", StringComparison.OrdinalIgnoreCase) || error.Contains("fields_differ", StringComparison.OrdinalIgnoreCase) || error.Contains("origem diferente", StringComparison.OrdinalIgnoreCase))
            return "Reconcilie uma vez. Se persistir, compare origem e cópia antes de qualquer remoção.";
        if (error.Contains("lote incompleto", StringComparison.OrdinalIgnoreCase) || error.Contains("única", StringComparison.OrdinalIgnoreCase) || error.Contains("amb", StringComparison.OrdinalIgnoreCase))
            return "Use Reconciliar pendências; se persistir, revise as cópias existentes sem repetir o lote.";
        if (error.Contains("durable", StringComparison.OrdinalIgnoreCase) || error.Contains("backup_missing", StringComparison.OrdinalIgnoreCase))
            return "PST ocupado: feche mensagens deste arquivo ou, na confirmação de remoção, autorize a validação com backup MSG íntegro retido. Preserve os backups.";
        return "Use Reconciliar pendências e Verificar pendentes; consulte o registro se persistir.";
    }
    internal const string RetainedMsgExplanation = "Se o Outlook mantiver o PST ocupado, permite validar o PST reaberto junto ao backup MSG íntegro. Preserve a pasta inteira; dois arquivos no mesmo disco não protegem contra falha do disco.";
    internal const string ExistingPstExplanation = "Novos arquivamentos preservam a estrutura: Caixa de Entrada → Inbox, conforme o idioma do PST. Subpastas mantêm seus nomes. O conteúdo existente será preservado; mensagens copiadas manualmente não entram na deduplicação automática.";
    internal const string CutoffExplanation = "Inclui todo o dia escolhido no fuso do Windows registrado no plano. Usa a data de recebimento; quando ela não está disponível, usa a data de envio. Rascunhos e tipos não suportados continuam pendentes.";
    internal static string CutoffLabel(DateOnly? cutoff, string? zone = null) => cutoff is { } value
        ? $"Até {value:dd/MM/yyyy}, inclusive" + (string.IsNullOrWhiteSpace(zone) ? "" : $" · fuso: {zone}")
        : "Todos os e-mails (sem data de corte)";
    internal static readonly JsonSerializerOptions Json = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };
    internal static string ValidateRoot(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) throw new ArgumentException("Escolha uma pasta local para guardar o arquivo e o histórico.");
        if (!System.IO.Path.IsPathFullyQualified(input)) throw new ArgumentException("Informe o caminho completo de uma pasta local.");
        string path = System.IO.Path.GetFullPath(input).TrimEnd(System.IO.Path.DirectorySeparatorChar);
        if (path.Length > 120) throw new ArgumentException("Escolha um caminho com até 120 caracteres.");
        if (path.StartsWith("\\\\", StringComparison.Ordinal) || !System.IO.Path.IsPathFullyQualified(path) || path.Length <= 3)
            throw new ArgumentException("Escolha uma pasta em um disco local, fora da raiz da unidade.");
        if (path.Split(System.IO.Path.DirectorySeparatorChar).Any(x => x.StartsWith("OneDrive", StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("Escolha uma pasta fora do OneDrive.");
        return path;
    }
    internal static string StateLabel(string state) => state switch
    {
        "Planned" => "Aguardando cópia", "Exported" => "Backup MSG salvo", "ImportStarted" => "Cópia em andamento",
        "Copied" => "Copiado; validação pendente", "Verified" => "Validado", "DeleteStarted" => "Movimentação em andamento",
        "BackupVerifiedSourceUnmanaged" => "Cópia no PST validada; sem ação na origem", "InDeletedItems" => "Original em Itens Excluídos", "PurgeStarted" => "Remoção em andamento", "Removed" => "Original removido",
        "SkippedByPolicy" => "Convites de reunião", "ReviewRequired" => "Precisa de revisão", "SkippedDraft" => "Rascunho pendente", "SkippedUnsupported" => "Tipo de item pendente",
        _ => state
    };
    internal static string? ValidateExistingPstPath(bool useExisting, string input)
    {
        if (!useExisting) return null;
        if (string.IsNullOrWhiteSpace(input)) throw new ArgumentException("Escolha o arquivo PST existente que receberá as cópias.");
        if (!Path.IsPathFullyQualified(input)) throw new ArgumentException("Informe o caminho completo do PST existente.");
        string path = Path.GetFullPath(input);
        if (path.StartsWith("\\\\", StringComparison.Ordinal) || path.Split(Path.DirectorySeparatorChar).Any(x => x.StartsWith("OneDrive", StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("Escolha um PST em um disco local, fora do OneDrive.");
        if (!string.Equals(Path.GetExtension(path), ".pst", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Selecione um arquivo com extensão .pst.");
        if (!File.Exists(path)) throw new ArgumentException("O PST escolhido não existe. Para criar um arquivo, selecione Criar novos arquivos PST.");
        return path;
    }
    internal static string JobSummary(IReadOnlyDictionary<string, int> states)
    {
        int total = states.Values.Sum();
        int verified = states.GetValueOrDefault("Verified"), removed = states.GetValueOrDefault("Removed"), trash = states.GetValueOrDefault("InDeletedItems");
        int sourceUnknown = states.GetValueOrDefault("BackupVerifiedSourceUnmanaged");
        int ignored = states.GetValueOrDefault("SkippedByPolicy");
        int pending = total - verified - removed - trash - sourceUnknown - ignored;
        var parts = new List<string> { $"{total:N0} itens" };
        if (ignored > 0) parts.Add($"{ignored:N0} convites de reunião");
        if (verified > 0) parts.Add($"{verified:N0} validados");
        if (removed > 0) parts.Add($"{removed:N0} originais removidos");
        if (sourceUnknown > 0) parts.Add($"{sourceUnknown:N0} cópias confirmadas; sem ação na origem");
        if (trash > 0) parts.Add($"{trash:N0} em Itens Excluídos");
        parts.Add($"{pending:N0} pendentes");
        return string.Join(" · ", parts);
    }
    internal static object PlanConfig(string directory, SourceSelection source, bool includeSubfolders, decimal volumeGiB, bool retainedMsgBackup, string? existingPstPath = null, DateOnly? cutoffDateInclusive = null, string? cutoffTimeZoneId = null, string? destinationPstPath = null, bool moveValidatedToDeletedItems=false)
        => new { RootDirectory = directory, SourceStoreId = source.StoreId, SourceFolderId = source.FolderId,
            SourceLabel = source.Label, IncludeSubfolders = includeSubfolders, TargetVolumeBytes = checked((long)volumeGiB * 1024 * 1024 * 1024),
            AllowRetainedMsgBackupForRemoval = retainedMsgBackup, ExistingPstPath = existingPstPath, DestinationPstPath=destinationPstPath, PreserveFolderStructure = true,
            MoveValidatedToDeletedItems=moveValidatedToDeletedItems,
            CutoffDateInclusive = cutoffDateInclusive, CutoffTimeZoneId = cutoffTimeZoneId ?? TimeZoneInfo.Local.Id };
}
