using System.Diagnostics;
using System.Text.Json;

namespace OutlookArchive.UI;

internal sealed partial class MainForm : Form
{
    private readonly EngineClient engine = new();
    private readonly EngineClient statusEngine = new();
    private readonly TextBox root = new() { Dock = DockStyle.Fill };
    private readonly ComboBox destinationMode = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260, AccessibleName = "Modo de destino PST" };
    private readonly TextBox existingPst = new() { Dock = DockStyle.Fill, PlaceholderText = "Selecione o arquivo PST existente…", AccessibleName = "Caminho do PST existente" };
    private readonly Button browsePst = Button("Escolher PST…");
    private readonly TextBox search = new() { Dock = DockStyle.Fill, PlaceholderText = "Filtrar pastas pelo nome…" };
    private readonly TreeView sources = new() { Dock = DockStyle.Fill, HideSelection = false, FullRowSelect = true };
    private readonly NumericUpDown volumeSize = new() { Minimum = 1, Maximum = 40, Value = 20, Width = 72 };
    private readonly CheckBox includeChildren = new() { Text = "Incluir subpastas", Checked = true, AutoSize = true };
    private readonly CheckBox cutoffEnabled = new() { Text = "Arquivar até (inclusive)", Checked = false, AutoSize = true, AccessibleName = "Limitar por data de corte" };
    private readonly DateTimePicker cutoffPicker = new() { Format = DateTimePickerFormat.Custom, CustomFormat = "dd/MM/yyyy", Value = DateTime.Today,
        Width = 155, Enabled = false, AccessibleName = "Data de corte inclusiva" };
    private readonly CheckBox retainedMsgBackup = new() { Text = "Permitir validação com backup MSG retido", Checked = false, AutoSize = true, AccessibleDescription = UiRules.RetainedMsgExplanation };
    private readonly ToolTip hints = new() { AutoPopDelay = 20000, InitialDelay = 300, ReshowDelay = 200 };
    private readonly ComboBox jobs = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label summary = new() { AutoSize = true, MaximumSize = new Size(690, 0), Text = "Selecione uma pasta de arquivo para consultar as execuções." };
    private readonly Label activity = new() { Dock = DockStyle.Fill, Text = "Pronto para começar", TextAlign = ContentAlignment.MiddleLeft };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Fill, Style = ProgressBarStyle.Continuous };
    private readonly DataGridView stateGrid = Grid();
    private readonly DataGridView volumeGrid = Grid();
    private readonly DataGridView errorGrid = Grid();
    private readonly Button connect = Button("Carregar caixas e pastas");
    private readonly Button plan = Button("Preparar arquivamento", true);
    private readonly Button copy = Button("Copiar / retomar", true);
    private readonly Button verify = Button("Verificar pendentes");
    private readonly Button reconcile = Button("Reconciliar pendências");
    private readonly Button pendingHelp = Button("Como resolver pendências");
    private readonly Button pendingReport = Button("Salvar relatório…");
    private JobStatus? displayedJob;
    private readonly Button previousPending = Button("Anterior");
    private readonly Button nextPending = Button("Próxima");
    private readonly Label pendingPageLabel = new() { AutoSize = true, Margin = new Padding(8) };
    private int pendingPage;
    private const int PendingPageSize = 200;
    private readonly Button remove = Button("Remover originais…");
    private readonly Button stop = Button("Interromper");
    private readonly Button openFolder = Button("Abrir pasta do arquivo");
    private readonly System.Windows.Forms.Timer timer = new() { Interval = 4000 };
    private List<SourceStore> sourceData = [];
    private bool busy, refreshing, loadingJobs;
    private string? activeRoot;
    private string? preparingJobId;
    internal MainForm(string? initialRoot = null, string? initialJob = null)
    {
        Text = "Arquivo de e-mails · Outlook clássico";
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleDimensions = new SizeF(96,96); AutoScaleMode=AutoScaleMode.Dpi;
        Font=new Font("Segoe UI",10); BackColor=Color.FromArgb(245,247,250);
        FormBorderStyle=FormBorderStyle.FixedSingle; MaximizeBox=false;
        var area=Screen.FromPoint(Cursor.Position).WorkingArea;
        Size=new Size(Math.Min(960,area.Width-30),Math.Min(760,area.Height-40));
        root.Text=initialRoot ?? Path.Combine(AppContext.BaseDirectory,"emailArchive");
        activeRoot=root.Text;
        engine.LogDirectory=statusEngine.LogDirectory=Path.Combine(root.Text,"logs");
        BuildWizard();
        using(var iconStream=typeof(MainForm).Assembly.GetManifestResourceStream("OutlookArchive.UI.ArchiveIcon"))
        {if(iconStream is not null){using var appIcon=new Icon(iconStream);Icon=(Icon)appIcon.Clone();}}
        connect.Click+=async (_,_)=>await Guarded("Carregando pastas do Outlook…",LoadSources);
        browsePst.Click+=(_,_)=>BrowseExistingPst();
        destinationMode.SelectedIndexChanged+=(_,_)=>UpdateActions();
        search.TextChanged+=(_,_)=>PopulateSources(); sources.AfterSelect+=(_,_)=>UpdateActions();
        jobs.SelectedIndexChanged+=(_,_)=>{if(!loadingJobs)ShowSelectedJob();};
        cutoffEnabled.CheckedChanged+=(_,_)=>UpdateActions();
        plan.Click+=async (_,_)=>await Guarded("Preparando o plano de arquivamento…",Plan);
        copy.Click+=async (_,_)=>{if(displayedJob?.Workflow is not null)await ContinueArchive();else await Execute("archive-run","Copiando as mensagens…");};
        verify.Click+=async (_,_)=>await Execute("archive-verify","Verificando as cópias pendentes…");
        reconcile.Click+=async (_,_)=>await Execute("archive-reconcile","Conferindo as pendências…");
        remove.Click+=async (_,_)=>await Remove();
        stop.Click+=(_,_)=>{
            if(MessageBox.Show(this,"Interromper esta operação? O histórico será preservado para retomada.","Interromper",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes)
            {engine.Interrupt();activity.Text="Interrupção solicitada; aguarde o encerramento.";}
        };
        previousPending.Click+=(_,_)=>{pendingPage--;ShowPendingPage();};
        nextPending.Click+=(_,_)=>{pendingPage++;ShowPendingPage();};
        pendingHelp.Click+=(_,_)=>MessageBox.Show(this,UiRules.PendingGuide,"Pendências",MessageBoxButtons.OK,MessageBoxIcon.Information);
        pendingReport.Click+=(_,_)=>SavePendingReport();
        timer.Tick+=async (_,_)=>{UpdateActions();if(busy && activeRoot is not null && !refreshing)await PollStatus();};
        FormClosing+=(_,e)=>{if(busy){e.Cancel=true;MessageBox.Show(this,"Use Interromper e aguarde antes de fechar.","Operação em andamento");}};
        Shown+=async (_,_)=>await Guarded("Preparando a pasta de trabalho…",async ()=>{
            Directory.CreateDirectory(root.Text);Directory.CreateDirectory(Path.Combine(root.Text,"logs"));Directory.CreateDirectory(Path.Combine(root.Text,"plans"));
            await RefreshJobs(initialJob);ShowStep(initialJob is null ? 0 : 5);activity.Text=initialJob is null ? "Escolha iniciar um backup ou continuar uma execução." : "Execução carregada. Use Continuar arquivamento para retomar os lotes.";
        });
        UpdateActions();
    }
    private static Panel InputHost(ComboBox input, bool stretch = true)
    {
        var host = new Panel { Dock = stretch ? DockStyle.Fill : DockStyle.None, Width = input.Width + 6, Margin = new Padding(0, 3, 0, 3) };
        input.Dock = DockStyle.Top; input.Margin = Padding.Empty; host.Controls.Add(input);
        void Fit()
        {
            int height = Math.Max(input.Height, input.Font.Height + 16);
            host.MinimumSize = new Size(0, height); host.Height = height;
        }
        input.FontChanged += (_, _) => Fit(); input.SizeChanged += (_, _) => Fit(); Fit();
        return host;
    }
    private static void BoundFlow(FlowLayoutPanel flow)
    {
        void Limit()
        {
            if (flow.Parent is not { } parent) return;
            // Iconic window dimensions must not become persistent wrapping limits.
            if (parent.FindForm()?.WindowState==FormWindowState.Minimized || parent.ClientSize.Width<=0) return;
            var limit = new Size(Math.Max(100, parent.ClientSize.Width - flow.Margin.Horizontal), 0);
            if (flow.MaximumSize != limit) flow.MaximumSize = limit;
        }
        if (flow.Parent is { } container) container.SizeChanged += (_, _) => Limit();
        Limit();
    }
    private static TableLayoutPanel NaturalTable(int columns, int rows)
    {
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = columns, RowCount = rows, Margin = Padding.Empty };
        if (columns == 1) panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (int i = 0; i < rows; i++) panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        return panel;
    }
    private static Label WrappedLabel(string text)
    {
        var label = new Label { Text = text, AutoSize = true, Dock = DockStyle.Top, ForeColor = Color.FromArgb(75, 89, 104), Margin = new Padding(0, 3, 0, 5) };
        label.ParentChanged += (_, _) => BoundLabel(label);
        return label;
    }
    private static void BoundLabel(Label label)
    {
        if (label.Parent is not { } parent) return;
        void Fit()
        {
            if (parent.FindForm()?.WindowState==FormWindowState.Minimized || parent.ClientSize.Width<=0) return;
            var maximum = new Size(Math.Max(100, parent.ClientSize.Width - label.Margin.Horizontal), 0);
            if (label.MaximumSize != maximum) label.MaximumSize = maximum;
        }
        parent.SizeChanged += (_, _) => Fit(); Fit();
    }
    private static Button Button(string text, bool primary = false)
    {
        var button = new Button
        {
            Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(10, 5, 10, 5),
            Margin = new Padding(0, 0, 7, 7), FlatStyle = FlatStyle.Flat,
            BackColor = primary ? Color.FromArgb(24, 85, 133) : Color.White,
            ForeColor = primary ? Color.White : Color.FromArgb(30, 51, 72), UseVisualStyleBackColor = false
        };
        void FitText()
        {
            Size textSize = TextRenderer.MeasureText(button.Text, button.Font);
            button.MinimumSize = new Size(textSize.Width + button.Padding.Horizontal + 10, textSize.Height + button.Padding.Vertical + 10);
        }
        button.FontChanged += (_, _) => FitText(); button.PaddingChanged += (_, _) => FitText();
        button.DpiChangedAfterParent += (_, _) => FitText(); FitText();
        button.EnabledChanged += (_, _) =>
        {
            // Native disabled text uses the Windows gray color. Keep its
            // background light instead of retaining a dark primary fill.
            button.BackColor = !button.Enabled ? Color.FromArgb(238,240,243) : primary ? Color.FromArgb(24, 85, 133) : Color.White;
            button.ForeColor = !button.Enabled ? Color.FromArgb(130,137,145) : primary ? Color.White : Color.FromArgb(30, 51, 72);
            button.FlatAppearance.BorderColor=button.Enabled ? Color.FromArgb(30,51,72) : Color.FromArgb(205,210,216);
        };
        return button;
    }
    private sealed class BoundedGrid : DataGridView
    {
        // AutoSize parents must size the viewport, never all rows in history.
        public override Size GetPreferredSize(Size proposedSize) => new(320, 180);
    }
    private static DataGridView Grid() => new BoundedGrid()
    {
        Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
        RowHeadersVisible = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
        BackgroundColor = Color.White, BorderStyle = BorderStyle.None, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        MultiSelect = false, AllowUserToResizeRows = false, AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders,
        ShowCellToolTips = false,
        ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
    };
    private void AppendLog(string line)
    {
        if (IsDisposed || !IsHandleCreated) return;
        if (InvokeRequired) { BeginInvoke((Action)(() => AppendLog(line))); return; }
        Directory.CreateDirectory(Path.Combine(root.Text,"logs"));
        File.AppendAllText(Path.Combine(root.Text,"logs","interface-"+DateTime.Today.ToString("yyyyMMdd")+".log"),DateTimeOffset.Now.ToString("O")+" "+line+Environment.NewLine);
        if(line.StartsWith("PLANEJAMENTO=",StringComparison.Ordinal) && Guid.TryParseExact(line[13..].Trim(),"N",out var preparing))preparingJobId=preparing.ToString("N");
        if(line.StartsWith("FLUXO=",StringComparison.Ordinal))
        {try{if(JsonSerializer.Deserialize<WorkflowInfo>(line[6..],UiRules.Json) is { } data)ApplyWorkflowProgress(data,true);}catch(JsonException){}return;}
        if(workflowRunning)return;
        if (busy && line.StartsWith("Planejamento:", StringComparison.Ordinal)) activity.Text = line;
        else if (busy && line.StartsWith("Validação:", StringComparison.Ordinal)) activity.Text = line;
        else if (busy && line.StartsWith("Cópia registrada:", StringComparison.Ordinal)) activity.Text = "Copiando e validando as mensagens…";
    }
    private async Task Guarded(string message, Func<Task> action)
    {
        if (busy) return;
        busy = true; activity.Text = message; progress.Style = ProgressBarStyle.Marquee; UpdateActions(); timer.Start();
        try { await action(); }
        catch (EngineInterruptedException)
        {
            try
            {
                await WaitForStatusPoll();
                if (File.Exists(Path.Combine(root.Text, "archive.sqlite")))await RefreshJobs(preparingJobId);
                if(displayedJob?.PlanningPending==true)ShowStep(5);
            }
            catch (Exception ex) { AppendLog("Não foi possível atualizar o histórico após a interrupção: " + ex.Message); }
            activity.Text = "Operação interrompida por você. Registros já gravados foram preservados.";
            AppendLog("Interrupção solicitada. Selecione a execução em Continuar backup; o planejamento retoma os checkpoints já gravados.");
        }
        catch (Exception ex)
        {
            bool planning=preparingJobId is not null;
            if(planning)
            {
                try{await WaitForStatusPoll();await RefreshJobs(preparingJobId);ShowStep(5);}
                catch(Exception refreshError){AppendLog("Não foi possível selecionar o planejamento salvo: "+refreshError.Message);}
            }
            string friendlyMessage=UiRules.ErrorMessage(ex.Message);
            activity.Text=planning?"Planejamento preservado. Após resolver o aviso, use Continuar arquivamento nesta execução.":"Operação não concluída. Consulte a pasta de logs para detalhes.";
            AppendLog(ex.ToString());MessageBox.Show(this,friendlyMessage,"Arquivamento",MessageBoxButtons.OK,MessageBoxIcon.Warning);
        }
        finally { preparingJobId=null;timer.Stop(); busy = false; progress.Style = ProgressBarStyle.Continuous; progress.Value = 0; if(displayedJob?.Workflow is { } flow)ApplyWorkflowProgress(flow,false); UpdateActions(); }
    }
    private async Task LoadSources()
    {
        var result = await engine.RunAsync(["archive-ui-sources"]);
        EnsureSuccess(result);
        var listing = JsonSerializer.Deserialize<SourceListing>(result.Output, UiRules.Json) ?? throw new InvalidDataException("Lista de pastas inválida.");
        sourceData = listing.Stores; PopulateSources();
        foreach (string warning in listing.Warnings) AppendLog(warning);
        activity.Text = $"{sourceData.Count} caixas carregadas. Selecione a origem na árvore.";
    }
    private void PopulateSources()
    {
        SourceSelection? previous = sources.SelectedNode?.Tag as SourceSelection;
        TreeNode? restore = null;
        sources.BeginUpdate(); sources.Nodes.Clear();
        foreach (var store in sourceData)
        {
            var top = new TreeNode(store.Name);
            foreach (var folder in store.Folders)
            {
                string path = string.IsNullOrWhiteSpace(folder.Path) ? folder.Name : folder.Path;
                if (search.Text.Length > 0 && !path.Contains(search.Text, StringComparison.CurrentCultureIgnoreCase) && !store.Name.Contains(search.Text, StringComparison.CurrentCultureIgnoreCase)) continue;
                var selected = new SourceSelection(store.StoreId, folder.FolderId, store.Name + " > " + path);
                var node = new TreeNode(path) { Tag = selected, ToolTipText = selected.Label }; top.Nodes.Add(node);
                if (previous?.FolderId == selected.FolderId && previous.StoreId == selected.StoreId) restore = node;
            }
            if (top.Nodes.Count > 0) sources.Nodes.Add(top);
        }
        sources.ShowNodeToolTips = true; sources.ExpandAll();
        if (sources.Nodes.Count > 0)
        {
            sources.SelectedNode = restore ?? sources.Nodes[0];
            sources.SelectedNode.EnsureVisible();
            sources.TopNode = restore ?? sources.Nodes[0];
        }
        sources.EndUpdate(); UpdateActions();
    }
    private void BrowseExistingPst()
    {
        using var dialog = new OpenFileDialog { Title = "Escolha um PST existente", Filter = "Arquivos de dados do Outlook (*.pst)|*.pst", CheckFileExists = true, Multiselect = false };
        if (File.Exists(existingPst.Text)) dialog.FileName = existingPst.Text;
        if (dialog.ShowDialog(this) == DialogResult.OK) existingPst.Text = dialog.FileName;
    }
    private async Task Plan()
    {
        if (sources.SelectedNode?.Tag is not SourceSelection source) throw new ArgumentException("Selecione uma pasta de origem.");
        string directory = UiRules.ValidateRoot(root.Text);
        string? selectedPst = UiRules.ValidateExistingPstPath(destinationMode.SelectedIndex == 1, existingPst.Text);
        DateOnly? cutoff = cutoffEnabled.Checked ? DateOnly.FromDateTime(cutoffPicker.Value) : null;
        string cutoffZone = TimeZoneInfo.Local.Id;
        string destination=ValidateDestination();
        string temporary=Path.Combine(directory,"plans","plan-"+Guid.NewGuid().ToString("N")+".json");
        try
        {
            var config = UiRules.PlanConfig(directory, source, includeChildren.Checked, volumeSize.Value, retainedMsgBackup.Checked, selectedPst, cutoff, cutoffZone, destination,autoMoveNew.Checked);
            await File.WriteAllTextAsync(temporary, JsonSerializer.Serialize(config, UiRules.Json));
            var result = await engine.RunAsync(["archive-plan", temporary], AppendLog); EnsureSuccess(result);
            string? id = result.Output.Split('\n').Select(x => x.Trim()).FirstOrDefault(x => x.StartsWith("EXECUCAO=", StringComparison.Ordinal))?[9..];
            activeRoot = directory;
            if(id is not null)EnsureSuccess(await engine.RunAsync(["archive-workflow-prepare",directory,id,autoMoveNew.Checked?"--move-validated-to-trash":"--keep-originals"],AppendLog));
            await RefreshJobs(id); ShowStep(5); activity.Text = "Plano preparado em lotes. Use Continuar arquivamento.";
        }
        finally { /* The reviewed plan remains in the local history. */ }
    }
    private async Task RefreshJobs(string? selectId = null)
    {
        string directory = UiRules.ValidateRoot(root.Text);
        if (!File.Exists(Path.Combine(directory, "archive.sqlite")))
        { activeRoot = directory; jobs.Items.Clear(); stateGrid.DataSource = null; volumeGrid.DataSource = null; errorGrid.DataSource = null; summary.Text = "Nenhuma execução nesta pasta. Selecione a origem e prepare o primeiro arquivamento."; activity.Text = "Pasta pronta para um novo arquivamento."; return; }
        var result = await statusEngine.RunAsync(["archive-status", directory]); EnsureSuccess(result);
        var data = JsonSerializer.Deserialize<List<JobStatus>>(result.Output, UiRules.Json) ?? [];
        string? current = selectId ?? (jobs.SelectedItem as JobStatus)?.Id;
        activeRoot = directory; loadingJobs = true;
        try
        {
            jobs.Items.Clear(); jobs.Items.AddRange(data.OrderByDescending(x => x.CreatedUtc).Cast<object>().ToArray());
            if (jobs.Items.Count > 0) jobs.SelectedItem = jobs.Items.Cast<JobStatus>().FirstOrDefault(x => x.Id == current) ?? jobs.Items.Cast<JobStatus>().FirstOrDefault(x=>x.PlanningPending||WizardRules.CanContinue(x.States)) ?? jobs.Items[0];
        }
        finally { loadingJobs = false; }
        ShowSelectedJob();
        activity.Text = "Histórico carregado.";
    }
    private async Task PollStatus()
    {
        refreshing = true;
        try
        {
            if (activeRoot is null || jobs.SelectedItem is not JobStatus selected || !File.Exists(Path.Combine(activeRoot, "archive.sqlite"))) return;
            string directory = activeRoot, selectedId = selected.Id;
            var result = await statusEngine.RunAsync(["archive-status", directory, selectedId]);
            if (result.ExitCode != 0) return;
            var data = JsonSerializer.Deserialize<List<JobStatus>>(result.Output, UiRules.Json) ?? [];
            if (string.Equals(activeRoot, directory, StringComparison.OrdinalIgnoreCase) && jobs.SelectedItem is JobStatus current &&
                current.Id == selectedId && data.FirstOrDefault(x => x.Id == selectedId) is { } updated)
                ShowJob(updated);
        }
        catch (Exception ex) when (ex is IOException or JsonException or InvalidOperationException) { /* A later poll retries read-only status. */ }
        finally { refreshing = false; }
    }
    private void ShowSelectedJob()
    {
        if (jobs.SelectedItem is JobStatus job) ShowJob(job);
        else { displayedJob = null; summary.Text = "Nenhuma execução selecionada."; stateGrid.DataSource = null; volumeGrid.DataSource = null; errorGrid.DataSource = null; }
        UpdateActions();
    }
    private void ShowJob(JobStatus job)
    {
        if (displayedJob?.Id != job.Id)
        {
            pendingPage = 0;loadingAutoMovement=true;
            try{autoMoveCurrent.Checked=RemovalPreferences.Read(root.Text,job.Id,Path.Combine(root.Text,"preferences","automatic-movement")) ?? job.Workflow?.MoveOriginals ?? job.MoveValidatedToDeletedItems;}
            finally{loadingAutoMovement=false;}
        }
        displayedJob = job;
        summary.Text = $"{job.SourceLabel ?? "Origem registrada no plano da execução"}\n{UiRules.CutoffLabel(job.CutoffDateInclusive, job.CutoffTimeZoneId)}\n{UiRules.JobSummary(job.States)}";
        if(job.PlanningPending)summary.Text=$"{job.SourceLabel??"Origem do planejamento salvo"}\n{UiRules.CutoffLabel(job.CutoffDateInclusive,job.CutoffTimeZoneId)}\nPlanejamento incompleto: {job.PlanningAnalyzed:N0}/{job.PlanningTotal:N0} itens analisados e salvos. Use Continuar arquivamento para retomar.";
        if (job.PreserveFolderStructure) summary.Text += "\nPastas: estrutura da origem, com as pastas padrão no idioma do PST.";
        if(job.Workflow is { } flow){summary.Text+=$"\nLotes: {flow.CompletedBatches}/{flow.TotalBatches} concluídos · {flow.ReviewBatches} com pendências";ApplyWorkflowProgress(flow,false);}
        stateGrid.DataSource = job.States.Select(x => new { Situação = UiRules.StateLabel(x.Key), Quantidade = x.Value }).ToList();
        volumeGrid.DataSource = job.Volumes.Select(x => new { Parte = x.Number, Tamanho = $"{x.ObservedBytes / 1024d / 1024:N1} MiB", Situação = x.State == "Ready" ? "Disponível" : "Precisa de revisão", Arquivo = x.PstPath }).ToList();
        RefreshPendingGroups();ShowPendingPage();
        if (errorGrid.Columns.Contains("Motivo")) errorGrid.Columns["Motivo"].FillWeight = 300;
        if (errorGrid.Columns.Contains("Orientação")) errorGrid.Columns["Orientação"].FillWeight = 300;
        if (volumeGrid.Columns.Contains("Arquivo")) volumeGrid.Columns["Arquivo"].FillWeight = 250;
    }
    private void ShowPendingPage()
    {
        var errors = FilteredPending();
        int pages = Math.Max(1, (errors.Count + PendingPageSize - 1) / PendingPageSize);
        pendingPage = Math.Clamp(pendingPage, 0, pages - 1);
        errorGrid.DataSource = errors.Skip(pendingPage * PendingPageSize).Take(PendingPageSize)
            .Select(x => new { Item = x.Key[..Math.Min(12, x.Key.Length)], Motivo = x.Error, Orientação = UiRules.PendingAction(x.Error) }).ToList();
        previousPending.Enabled = pendingPage > 0;
        nextPending.Enabled = pendingPage + 1 < pages;
        pendingPageLabel.Text = $"Página {pendingPage + 1} de {pages} · {errors.Count:N0} pendências";
    }
    private void SavePendingReport()
    {
        if (displayedJob is not { } job) return;
        try { string folder=Path.Combine(root.Text,"reports");Directory.CreateDirectory(folder);string file=Path.Combine(folder,$"pendencias-{job.Id[..8]}-{DateTime.Now:yyyyMMdd-HHmmss}.csv"); File.WriteAllText(file, UiRules.PendingReport(job), new System.Text.UTF8Encoding(true)); activity.Text = "Relatório salvo em emailArchive\\reports."; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { MessageBox.Show(this, ex.Message, "Não foi possível salvar o relatório", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
    private async Task Execute(string command, string message)
    {
        if (jobs.SelectedItem is not JobStatus job || activeRoot is null) return;
        string directory = activeRoot;
        await Guarded(message, async () =>
        {
            string[] arguments = command == "archive-verify"
                ? UiRules.VerificationArguments(directory, job.Id, RemovalPreferences.Read(directory, job.Id) == true) : [command, directory, job.Id];
            if(command is "archive-run" or "archive-workflow-run" && autoMoveCurrent.Checked && RemovalPreferences.Read(directory,job.Id)==true)
                arguments=[..arguments,"--allow-retained-msg-backup"];
            if(command is "archive-run" or "archive-verify" or "archive-workflow-run")
                arguments=[..arguments,autoMoveCurrent.Checked ? "--move-validated-to-trash" : "--keep-originals"];
            if(command=="archive-workflow-run")arguments=[..arguments,"--reconcile-pending"];
            var result = await engine.RunAsync(arguments, AppendLog);
            if (result.WasInterrupted) throw new EngineInterruptedException();
            await WaitForStatusPoll(); await RefreshJobs(job.Id);
            if (result.ExitCode == 3) activity.Text = "Operação encerrada com pendências. Confira as contagens e a aba Pendências.";
            else if(result.ExitCode==4)activity.Text="Arquivamento pausado. Os lotes concluídos foram preservados; use Continuar arquivamento para retomar.";
            else if (result.ExitCode == 0) activity.Text = "Operação concluída. Confira o resultado da execução.";
            else { activity.Text = "Operação interrompida; o histórico foi preservado para retomada."; if (result.Error.Length > 0) MessageBox.Show(this, result.Error, "Operação interrompida", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            if(command=="archive-workflow-run" && result.ExitCode is 0 or 3 && displayedJob is { } completed)ShowWorkflowSummary(completed);
            if (NewItemsFollowUp.ShouldCheck(command, result)) await CheckNewItems(directory, job.Id);
        });
    }
    private async Task CheckNewItems(string directory, string previousJob)
    {
        string completedActivity = activity.Text;
        activity.Text = "Conferindo se há novos e-mails fora do plano…";
        var followUp = await NewItemsFollowUp.ExecuteAsync(directory, previousJob,
            arguments => engine.RunAsync(arguments, arguments[0] == "archive-check-new" ? null : AppendLog),
            count => Task.FromResult(MessageBox.Show(this,
                $"Há {count:N0} novos e-mails fora do escopo desta execução. Deseja preparar outro plano para eles?\n\nO novo plano mantém o mesmo recorte de pastas, data de corte e destino. A execução atual será preservada.\n\nPreparar o plano não inicia a cópia. Você poderá conferi-lo e depois usar Continuar arquivamento.",
                "Novos e-mails encontrados", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes));
        if (followUp.NewJobId is not null)
        {
            try
            {
                await WaitForStatusPoll(); await RefreshJobs(followUp.NewJobId);
                activity.Text = "Novo plano preparado. Confira o resultado e use Continuar arquivamento quando desejar.";
                AppendLog($"Novo plano preparado: {followUp.NewJobId}. A cópia não foi iniciada.");
            }
            catch (Exception ex)
            {
                AppendLog($"O novo plano {followUp.NewJobId} foi registrado, mas o histórico não pôde ser atualizado: {ex.Message}");
                activity.Text = "Novo plano registrado; volte ao início para selecioná-lo. A cópia não foi iniciada.";
            }
        }
        else if (followUp.WasInterrupted)
        {
            activity.Text = "Consulta ou novo plano interrompido. A execução anterior foi preservada.";
            AppendLog("A etapa opcional de novos e-mails foi interrompida. Consulte o histórico para verificar se outro plano chegou a ser registrado.");
        }
        else if (followUp.Warning is not null)
        {
            activity.Text = "Execução anterior preservada. Há um aviso na consulta dos novos e-mails; consulte os logs.";
            AppendLog(followUp.Warning);
        }
        else if (followUp.NewItemsCount > 0)
        {
            activity.Text = $"Há {followUp.NewItemsCount:N0} novos e-mails fora do plano. A execução atual foi preservada.";
            AppendLog($"Novo plano não solicitado para os {followUp.NewItemsCount:N0} e-mails encontrados fora do escopo atual.");
        }
        else
        {
            activity.Text = completedActivity;
            AppendLog("Não há novos e-mails fora do escopo desta execução.");
        }
    }
    private async Task WaitForStatusPoll() { while (refreshing) await Task.Delay(80); }
    private async Task Remove()
    {
        if (busy || jobs.SelectedItem is not JobStatus selected || activeRoot is null) return;
        string directory = activeRoot; string id = selected.Id;
        await Guarded("Conferindo execução antes da remoção…", async () =>
        {
            await RefreshJobs(id);
            if (jobs.SelectedItem is not JobStatus current || current.Id != id) throw new InvalidDataException("Execução não encontrada.");
            using var dialog = new RemovalDialog(current, directory);
            if (dialog.ShowDialog(this) != DialogResult.OK) { activity.Text = "Remoção cancelada. Originais preservados."; return; }
            string[] validationArguments = UiRules.VerificationArguments(directory, id, dialog.AllowRetainedBackup);
            string[] removalArguments = dialog.AllowRetainedBackup ? ["archive-remove", directory, id, dialog.Mode, "--allow-retained-msg-backup"] : ["archive-remove", directory, id, dialog.Mode];
            if (current.States.GetValueOrDefault("Copied") > 0)
            {
                activity.Text = "Validando somente as cópias pendentes antes da remoção…";
                var validation = await engine.RunAsync(validationArguments, AppendLog);
                if (validation.WasInterrupted) throw new EngineInterruptedException();
                if (validation.ExitCode is not (0 or 3)) { EnsureSuccess(validation); return; }
                await WaitForStatusPoll(); await RefreshJobs(id);
            }
            if (jobs.SelectedItem is not JobStatus validated || !validated.States.Any(x => x.Value > 0 && x.Key is "Verified" or "InDeletedItems" or "DeleteStarted" or "PurgeStarted"))
            {
                activity.Text = "Nenhum original foi removido: ainda não há mensagens aprovadas para remoção.";
                AppendLog(activity.Text);
                foreach (var reason in (displayedJob?.Errors ?? []).GroupBy(x => x.Error).OrderByDescending(x => x.Count()).Take(5))
                    AppendLog($"Bloqueio de remoção — {reason.Count():N0} itens: {reason.Key}");
                MessageBox.Show(this, "Nenhum original passou em todas as verificações exigidas. Abra Pendências ou salve o relatório para consultar os motivos. Se o único bloqueio for o PST ocupado, a confirmação oferece a opção explícita de validar também com backup MSG retido.", "Remoção bloqueada pela validação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            activity.Text = "Revalidando e removendo somente os originais elegíveis…";
            var result = await engine.RunAsync(removalArguments, AppendLog);
            if (result.WasInterrupted) throw new EngineInterruptedException();
            await WaitForStatusPoll(); await RefreshJobs(id);
            activity.Text = result.ExitCode == 0 ? "Etapa de remoção concluída. Confira os estados finais." : "Remoção encerrada com pendências; confira as pendências. Itens sem validação permanecem protegidos.";
            if (result.ExitCode is not (0 or 3) && result.Error.Length > 0) MessageBox.Show(this, result.Error, "Remoção não concluída", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        });
    }
    private void UpdateActions()
    {
        bool job = jobs.SelectedItem is JobStatus && activeRoot is not null;
        bool ready=job && displayedJob?.PlanningPending!=true;
        connect.Enabled = !busy;
        destinationMode.Enabled = !busy; existingPst.Enabled = browsePst.Enabled = !busy && destinationMode.SelectedIndex == 1;
        root.ReadOnly = busy; search.Enabled = sources.Enabled = includeChildren.Enabled = retainedMsgBackup.Enabled = volumeSize.Enabled = jobs.Enabled = !busy;
        cutoffEnabled.Enabled = !busy; cutoffPicker.Enabled = !busy && cutoffEnabled.Checked;
        plan.Enabled = !busy && sources.SelectedNode?.Tag is SourceSelection;
        var states=displayedJob?.States ?? new Dictionary<string,int>();
        copy.Enabled=!busy && ready && WizardRules.CanCopy(states);
        verify.Enabled=!busy && ready && WizardRules.CanVerify(states,autoMoveCurrent.Checked);
        reconcile.Enabled=!busy && ready && WizardRules.CanReconcile(states);
        pendingReport.Enabled = !busy && job;
        remove.Enabled = !busy && ready && WizardRules.CanRemove(states);
        autoMoveNew.Enabled=autoMoveCurrent.Enabled=!busy;
        continueArchive.Enabled=!busy && job && (displayedJob?.PlanningPending==true || WizardRules.NextAction(states,autoMoveCurrent.Checked)!=ContinueAction.None);
        pauseAfterBatch.Enabled=busy && workflowRunning && engine.IsRunning && !batchPauseRequested;
        stop.Enabled = busy && engine.IsRunning;
        UpdateWizardActions();
    }
    private static void EnsureSuccess(EngineResult result)
    {
        if (result.WasInterrupted) throw new EngineInterruptedException();
        if (result.ExitCode != 0) throw new InvalidOperationException(string.IsNullOrWhiteSpace(result.Error) ? $"O motor encerrou com código {result.ExitCode}. Confira o registro da operação." : result.Error.Trim());
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing) { timer.Dispose(); hints.Dispose(); }
        base.Dispose(disposing);
    }
}

internal sealed class RemovalDialog : Form
{
    private readonly CheckBox allowRetainedBackup = new() { Text = "Permitir backup MSG íntegro retido quando o PST estiver ocupado", AutoSize = true, Dock = DockStyle.Top };
    internal bool AllowRetainedBackup => allowRetainedBackup.Checked;
    private readonly ComboBox modes = new() { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
    private readonly TextBox confirmation = new() { Dock = DockStyle.Fill, PlaceholderText = "Digite REMOVER para habilitar a ação" };
    internal string Mode => modes.SelectedIndex == 1 ? "mailbox" : "deleted-items";
    internal RemovalDialog(JobStatus job, string root)
    {
        Text = "Remover originais após validação"; Font = new Font("Segoe UI", 10);
        var available = Screen.FromPoint(Cursor.Position).WorkingArea;
        Size = new Size(Math.Min(740, available.Width - 30), Math.Min(820, available.Height - 40));
        FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = MinimizeBox = false; StartPosition = FormStartPosition.CenterParent;
        var panel = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(22), ColumnCount = 1, RowCount = 7,
            AutoScroll = true, AutoScrollMinSize = new Size(0, 765) };
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 140)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 110)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        string pstDetail = (job.DestinationPstPath ?? job.ExistingPstPath) is { } destination ? "PST inicial: "+destination : "PSTs conforme o plano registrado.";
        panel.Controls.Add(new TextBox { Text = $"Origem: {job.SourceLabel ?? "conforme plano registrado"}\r\nExecução: {job.Id}\r\nHistórico: {root}\r\n{pstDetail}\r\nPeríodo: {UiRules.CutoffLabel(job.CutoffDateInclusive, job.CutoffTimeZoneId)}\r\n\r\nValidados: {job.States.GetValueOrDefault("Verified")} · Em Itens Excluídos: {job.States.GetValueOrDefault("InDeletedItems")}", Dock = DockStyle.Fill, ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.White });
        modes.Items.AddRange(["Mover para Itens Excluídos", "Remover da caixa (também de Itens Excluídos)"]); modes.SelectedIndex = 0; panel.Controls.Add(modes);
        string validationMode = job.RetainedMsgBackupForRemoval ? "Este plano permite validação com backup MSG retido. " + UiRules.RetainedMsgExplanation : "Este plano exige leitura exclusiva do PST para validar a persistência.";
        var explanation = new Label { Dock = DockStyle.Fill, AutoSize = true, Text = $"Há {job.States.GetValueOrDefault("Copied"):N0} cópias aguardando validação. Ao confirmar, o motor validará essas cópias e removerá apenas os originais aprovados. Pendências de outros itens não impedem a remoção parcial; itens sem validação permanecem preservados.\n\n" + validationMode + "\n\nMover para Itens Excluídos não libera a quota da caixa. Remover da caixa pode ficar sujeito à retenção do Microsoft 365 e não comprova quota liberada." };
        void FitExplanation()
        {
            var maximum = new Size(Math.Max(200, panel.ClientSize.Width - panel.Padding.Horizontal - explanation.Margin.Horizontal - SystemInformation.VerticalScrollBarWidth), 0);
            if (explanation.MaximumSize != maximum) explanation.MaximumSize = maximum;
        }
        panel.SizeChanged += (_, _) => FitExplanation(); FitExplanation(); panel.Controls.Add(explanation);
        panel.Controls.Add(new Label { Text = "Confirmação para esta execução:", AutoSize = true }); panel.Controls.Add(confirmation);
        var backupChoice = new Panel { Dock = DockStyle.Fill };
        allowRetainedBackup.Checked = job.RetainedMsgBackupForRemoval || RemovalPreferences.Read(root, job.Id) == true;
        backupChoice.Controls.Add(new Label { Text = "Mantenha os backups MSG após a operação. A opção acima também exige reabrir o PST e conferir o conteúdo de cada mensagem.", Dock = DockStyle.Bottom, Height = 52, ForeColor = Color.FromArgb(100, 71, 23) });
        backupChoice.Controls.Add(allowRetainedBackup); panel.Controls.Add(backupChoice);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
        var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel, AutoSize = true };
        var apply = new Button { Text = "Confirmar remoção", DialogResult = DialogResult.OK, AutoSize = true, Enabled = false };
        confirmation.TextChanged += (_, _) => apply.Enabled = confirmation.Text == "REMOVER";
        apply.Click += (_, _) =>
        {
            bool blockedByBusyPst = job.Errors.Any(x => x.Error.Contains("durable", StringComparison.OrdinalIgnoreCase) || x.Error.Contains("volume-reopen", StringComparison.OrdinalIgnoreCase));
            if (!job.RetainedMsgBackupForRemoval && !allowRetainedBackup.Checked && blockedByBusyPst)
            {
                if (MessageBox.Show(this, "A última tentativa foi bloqueada porque o Outlook mantém o PST ocupado.\n\nDeseja permitir a validação do PST reaberto junto ao backup MSG íntegro retido? Cada mensagem continuará sendo conferida e os backups deverão ser preservados.", "Como validar o PST ocupado", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                { DialogResult = DialogResult.None; return; }
                allowRetainedBackup.Checked = true;
            }
            try { RemovalPreferences.Save(root, job.Id, allowRetainedBackup.Checked); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            { MessageBox.Show(this, "A escolha será usada nesta operação, mas não pôde ser lembrada: " + ex.Message); }
        };
        buttons.Controls.Add(cancel); buttons.Controls.Add(apply); panel.Controls.Add(buttons); Controls.Add(panel); CancelButton = cancel;
    }
}
