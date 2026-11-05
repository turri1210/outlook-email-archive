using System.Text.Json;

namespace OutlookArchive.UI;

internal static class UiTests
{
    internal static int Run()
    {
        int assertions = 0;
        void Check(bool condition, string message) { assertions++; if (!condition) throw new InvalidOperationException(message); }
        void Reject(string path)
        {
            try { UiRules.ValidateRoot(path); throw new InvalidOperationException("Caminho inválido aceito: " + path); }
            catch (ArgumentException) { assertions++; }
        }
        try
        {
            NewItemsTests(Check);
            Check(UiRules.ErrorMessage("MissingMemberException: Could not get dispid for Count (error: 0x800706BA).").StartsWith("Abra o Outlook clássico"),"Falha de conexão não orienta a abrir o Outlook clássico.");
            Check(UiRules.ErrorMessage("COMException: 0x800401e3").Contains("O novo Outlook não é compatível"),"Outlook fechado não recebe orientação correta.");
            Check(UiRules.ErrorMessage("PST indisponível")=="PST indisponível","Erro de PST foi confundido com Outlook fechado.");
            WizardActionTests(Check);
            Check(WizardRules.CanVerify(new Dictionary<string,int>{{"Verified",5}},true),"Movimento automático não habilitou itens já validados.");
            Check(!WizardRules.CanVerify(new Dictionary<string,int>{{"InDeletedItems",5}},true),"Itens já movidos habilitaram repetição automática.");
            Check(WizardRules.CanProcess(new Dictionary<string,int>{{"Copied",30}},true),"Fluxo não aceita cópias pendentes.");
            Check(!WizardRules.CanProcess(new Dictionary<string,int>{{"InDeletedItems",30}},true),"Fluxo habilitou lote já concluído.");
            foreach(string state in new[]{"Removed","SkippedByPolicy","BackupVerifiedSourceUnmanaged","SkippedDraft","SkippedUnsupported"})
            {var values=new Dictionary<string,int>{{state,1}};Check(!WizardRules.CanCopy(values)&&!WizardRules.CanVerify(values)&&!WizardRules.CanReconcile(values)&&!WizardRules.CanRemove(values),"Ação sem trabalho disponível: "+state);}
            Check(WizardRules.CanCopy(new Dictionary<string,int>{{"Planned",1}}),"Plano não habilita cópia.");
            Check(WizardRules.CanVerify(new Dictionary<string,int>{{"Copied",1}}),"Cópia não habilita verificação.");
            Check(!WizardRules.CanRemove(new Dictionary<string,int>{{"Copied",1}}),"Cópias não validadas habilitam remoção.");
            Check(WizardRules.CanRemove(new Dictionary<string,int>{{"Verified",1}}),"Validados não habilitam remoção.");
            Check(WizardRules.CanReconcile(new Dictionary<string,int>{{"ReviewRequired",1}}),"Revisão não habilita reconciliação.");
            Check(UiRules.VerificationArguments("C:\\Arquivo", "job", false).SequenceEqual(new[] { "archive-verify", "C:\\Arquivo", "job", "--pending-only" }), "Validação da interface não usa somente pendentes.");
            Check(UiRules.VerificationArguments("C:\\Arquivo", "job", true).SequenceEqual(new[] { "archive-verify", "C:\\Arquivo", "job", "--pending-only", "--allow-retained-msg-backup" }), "Validação incremental perdeu a autorização de backup.");
            string incompleteEngine = Path.Combine(Path.GetTempPath(), "outlook-incomplete-engine-" + Guid.NewGuid().ToString("N"), "OutlookArchive.dll");
            Directory.CreateDirectory(Path.GetDirectoryName(incompleteEngine)!);
            try
            {
                File.WriteAllText(incompleteEngine, "Synthetic fixture; never execute.");
                var incomplete = new EngineClient(incompleteEngine);
                try { incomplete.RunAsync(["archive-remove", "C:\\Arquivo", "synthetic", "mailbox"]).GetAwaiter().GetResult(); throw new Exception("Instalação incompleta iniciou uma operação."); }
                catch (InvalidOperationException ex) { Check(ex.Message.Contains("Instalação incompleta") && ex.Message.Contains("native-identity.exe") && !incomplete.IsRunning, "Componentes ausentes não foram detectados antes da operação."); }
            }
            finally { File.Delete(incompleteEngine); Directory.Delete(Path.GetDirectoryName(incompleteEngine)!); }
            string optionDirectory = Path.Combine(Path.GetTempPath(), "outlook-removal-options-" + Guid.NewGuid().ToString("N"));
            try
            {
                Check(RemovalPreferences.Read("C:\\Arquivo", "job-a", optionDirectory) is null, "Autorização criada sem confirmação.");
                RemovalPreferences.Save("C:\\Arquivo", "job-a", true, optionDirectory);
                Check(RemovalPreferences.Read("c:\\arquivo", "job-a", optionDirectory) == true, "Escolha confirmada não foi lembrada.");
                Check(RemovalPreferences.Read("C:\\Arquivo", "job-b", optionDirectory) is null, "Escolha vazou para outra execução.");
                Check(RemovalPreferences.Read("C:\\Outro", "job-a", optionDirectory) is null, "Escolha vazou para outro histórico.");
                RemovalPreferences.Save("C:\\Arquivo", "job-a", false, optionDirectory);
                Check(RemovalPreferences.Read("C:\\Arquivo", "job-a", optionDirectory) == false, "Revogação não foi lembrada.");
            }
            finally { if (Directory.Exists(optionDirectory)) Directory.Delete(optionDirectory, true); }
            Check(UiRules.CanPrepareRemoval(new Dictionary<string, int> { ["Copied"] = 5, ["ReviewRequired"] = 2 }), "Pendências bloquearam a preparação da remoção parcial.");
            Check(UiRules.CanPrepareRemoval(new Dictionary<string, int> { ["Verified"] = 5, ["ReviewRequired"] = 2 }), "Pendências bloquearam itens validados.");
            Check(!UiRules.CanPrepareRemoval(new Dictionary<string, int> { ["ReviewRequired"] = 2, ["SkippedUnsupported"] = 1 }), "Revisões sem cópia habilitaram remoção.");
            Check(UiRules.ValidateRoot("C:\\ArquivoOutlook") == "C:\\ArquivoOutlook", "Raiz local não aceita.");
            Reject(""); Reject("relativo"); Reject("C:\\"); Reject("\\\\server\\share\\arquivo"); Reject("C:\\Users\\Usuario\\OneDrive - Empresa\\arquivo"); Reject("C:\\" + new string('a', 121));
            Check(UiRules.StateLabel("Verified") == "Validado", "Estado validado incorreto.");
            Check(UiRules.JobSummary(new Dictionary<string,int> { ["SkippedByPolicy"]=38 }).Contains("0 pendentes"), "Reuniões desconsideradas não devem ser pendências.");
            Check(UiRules.StateLabel("SkippedByPolicy") == "Convites de reunião", "Rótulo de item desconsiderado ausente.");
            Check(UiRules.StateLabel("BackupVerifiedSourceUnmanaged").Contains("sem ação na origem"), "Cópia sem origem não deve ser declarada excluída.");
            Check(UiRules.JobSummary(new Dictionary<string,int> { ["BackupVerifiedSourceUnmanaged"]=2 }).Contains("0 pendentes"), "Cópias confirmadas sem origem não são pendências.");
            Check(UiRules.StateLabel("Copied").Contains("pendente"), "Cópia não deve aparecer como validada.");
            Check(UiRules.StateLabel("SkippedDraft") == "Rascunho pendente", "Rascunho não deve aparecer como concluído.");
            Check(UiRules.PendingAction("volume-reopen-pending: busy").Contains("Verificar pendentes"), "PST ocupado não orienta a verificação incremental.");
            Check(UiRules.PendingAction("Informações relevantes do lote não correspondem").Contains("antes de qualquer remoção"), "Divergência permite orientação insegura.");
            string report = UiRules.PendingReport(new JobStatus { Id = "test", States = new() { ["Copied"] = 4 }, Errors = [new() { Key = "abcdef", Error = "=danger;\"texto\"\nlinha" }] });
            Check(report.Contains("antes de remover originais") && report.Contains("\"'=danger;\"\"texto\"\"\nlinha\""), "Relatório perde orientação, escapes ou proteção de fórmulas.");
            Check(UiRules.JobSummary(new Dictionary<string, int> { ["Removed"] = 5 }) == "5 itens · 5 originais removidos · 0 pendentes", "Resumo confunde remoção concluída com validação pendente.");
            Check(UiRules.JobSummary(new Dictionary<string, int> { ["Verified"] = 3, ["SkippedDraft"] = 1, ["SkippedUnsupported"] = 1 }).EndsWith("2 pendentes"), "Resumo não inclui todos os itens pendentes.");
            string json = "[{\"Id\":\"0123456789abcdef0123456789abcdef\",\"CreatedUtc\":\"2026-10-26T12:00:00Z\",\"SourceLabel\":\"Caixa > Recebidos\",\"States\":{\"Verified\":3,\"SkippedDraft\":1},\"Volumes\":[]}]";
            var jobs = JsonSerializer.Deserialize<List<JobStatus>>(json, UiRules.Json)!;
            CancellationTests(Check);
            Check(jobs[0].States.Values.Sum() == 4 && jobs[0].SourceLabel == "Caixa > Recebidos", "Contrato do histórico inválido.");
            using var form = new MainForm();
            form.CreateControl();
            var controls = Descendants(form).ToList();
            Check(controls.OfType<TreeView>().Count() == 1, "Árvore de origem ausente.");
            Check(controls.OfType<DataGridView>().Count() == 3, "Grades de estados, volumes e pendências ausentes.");
            Check(controls.OfType<DataGridView>().All(x => !x.ShowCellToolTips), "As grades não devem exibir tooltips automáticas sobre os estados ou conteúdos.");
            Check(!controls.OfType<Button>().Single(x => x.Text == "Salvar relatório…").Enabled, "Relatório habilitado sem execução.");
            Check(controls.OfType<Button>().Single(x => x.Text == "Remover originais…").Enabled == false, "Remoção habilitada sem execução.");
            Check(controls.OfType<Button>().Single(x => x.Text == "Copiar / retomar").Enabled == false, "Cópia habilitada sem execução.");
            Check(!controls.OfType<CheckBox>().Single(x => x.Text == "Permitir validação com backup MSG retido").Checked, "Prova alternativa habilitada por padrão.");
            var cutoffEnabled = controls.OfType<CheckBox>().Single(x => x.AccessibleName == "Limitar por data de corte");
            var cutoffPicker = controls.OfType<DateTimePicker>().Single(x => x.AccessibleName == "Data de corte inclusiva");
            Check(!cutoffEnabled.Checked && !cutoffPicker.Enabled, "Data de corte habilitada por padrão.");
            cutoffEnabled.Checked = true;
            Check(cutoffPicker.Enabled && cutoffPicker.CustomFormat == "dd/MM/yyyy", "Data de corte não habilita o calendário brasileiro.");
            cutoffEnabled.Checked = false;
            Check(!cutoffPicker.Enabled, "Desmarcar o corte não desabilita o calendário.");
            Check(UiRules.CutoffLabel(new DateOnly(2024, 12, 31)).Contains("31/12/2024, inclusive"), "Resumo não informa corte inclusivo.");
            Check(UiRules.CutoffLabel(null).Contains("sem data de corte"), "Plano antigo não aparece sem corte.");
            var destinationMode = controls.OfType<ComboBox>().Single(x => x.AccessibleName == "Modo de destino PST");
            var existingPst = controls.OfType<TextBox>().Single(x => x.AccessibleName == "Caminho do PST existente");
            Check(destinationMode.SelectedIndex == 0 && !existingPst.Enabled, "Modo padrão não é criar novos PSTs.");
            destinationMode.SelectedIndex = 1; Check(existingPst.Enabled, "Campo do PST existente não foi habilitado."); destinationMode.SelectedIndex = 0;
            Check(UiRules.ValidateExistingPstPath(false, "arquivo-inválido") is null, "Modo novo considera caminho de PST anterior.");
            void RejectPst(string path)
            {
                try { UiRules.ValidateExistingPstPath(true, path); throw new InvalidOperationException("PST inválido aceito: " + path); }
                catch (ArgumentException) { assertions++; }
            }
            RejectPst(""); RejectPst("relativo.pst"); RejectPst("\\\\server\\share\\arquivo.pst"); RejectPst("C:\\Users\\Usuario\\OneDrive\\arquivo.pst");
            string temporary = Path.Combine(Path.GetTempPath(), "outlook-ui-test-" + Guid.NewGuid().ToString("N") + ".pst");
            string notPst = Path.ChangeExtension(temporary, ".ost");
            try
            {
                RejectPst(temporary); File.WriteAllText(temporary, "Fixture de caminho; não é um PST Outlook."); File.WriteAllText(notPst, "Fixture de extensão.");
                RejectPst(notPst);
                Check(UiRules.ValidateExistingPstPath(true, temporary) == temporary, "Caminho PST existente não aceito.");
                using var existingConfig = JsonDocument.Parse(JsonSerializer.Serialize(UiRules.PlanConfig("C:\\ArquivoOutlook", new SourceSelection("store", "folder", "Caixa"), true, 20, false, temporary), UiRules.Json));
                Check(existingConfig.RootElement.GetProperty("ExistingPstPath").GetString() == temporary, "PST existente não preservado na configuração.");
                Check(existingConfig.RootElement.GetProperty("PreserveFolderStructure").GetBoolean(), "Novo plano não preserva estrutura de pastas.");
                using var newConfig = JsonDocument.Parse(JsonSerializer.Serialize(UiRules.PlanConfig("C:\\ArquivoOutlook", new SourceSelection("store", "folder", "Caixa"), true, 20, false), UiRules.Json));
                Check(newConfig.RootElement.GetProperty("ExistingPstPath").ValueKind == JsonValueKind.Null, "Modo novo seleciona PST existente.");
            }
            finally { File.Delete(temporary); File.Delete(notPst); }
            var source = new SourceSelection("store", "folder", "Caixa > Recebidos");
            using (var dated = JsonDocument.Parse(JsonSerializer.Serialize(UiRules.PlanConfig("C:\\ArquivoOutlook", source, true, 20, false,
                cutoffDateInclusive: new DateOnly(2024, 12, 31), cutoffTimeZoneId: "E. South America Standard Time"), UiRules.Json)))
            {
                Check(dated.RootElement.GetProperty("CutoffDateInclusive").GetString() == "2024-12-31", "Data de corte não preservada em formato ISO.");
                Check(dated.RootElement.GetProperty("CutoffTimeZoneId").GetString() == "E. South America Standard Time", "Fuso do corte não preservado.");
            }
            using (var unrestricted = JsonDocument.Parse(JsonSerializer.Serialize(UiRules.PlanConfig("C:\\ArquivoOutlook", source, true, 20, false), UiRules.Json)))
                Check(unrestricted.RootElement.GetProperty("CutoffDateInclusive").ValueKind == JsonValueKind.Null, "Plano padrão aplica um corte indevido.");
            var datedStatus = JsonSerializer.Deserialize<JobStatus>("{\"Id\":\"0123456789abcdef0123456789abcdef\",\"CutoffDateInclusive\":\"2024-12-31\",\"CutoffTimeZoneId\":\"UTC\"}", UiRules.Json)!;
            Check(datedStatus.CutoffDateInclusive == new DateOnly(2024, 12, 31) && datedStatus.CutoffTimeZoneId == "UTC", "Contrato de status não preserva o corte.");
            Check(jobs[0].CutoffDateInclusive is null, "Histórico antigo ganhou data de corte.");
            foreach (bool retained in new[] { false, true })
            {
                using var document = JsonDocument.Parse(JsonSerializer.Serialize(UiRules.PlanConfig("C:\\ArquivoOutlook", source, true, 20, retained), UiRules.Json));
                Check(document.RootElement.GetProperty("AllowRetainedMsgBackupForRemoval").GetBoolean() == retained, "Escolha avançada não preservada na configuração.");
            }
            using var removal = new RemovalDialog(jobs[0], "C:\\ArquivoOutlook");
            var removeControls = Descendants(removal).ToList();
            var confirm = removeControls.OfType<Button>().Single(x => x.Text == "Confirmar remoção");
            Check(!confirm.Enabled && removal.Mode == "deleted-items", "Remoção não exige confirmação ou modo padrão incorreto.");
            Check(!removal.AllowRetainedBackup, "Alternativa de backup MSG habilitada sem autorização.");
            removal.CreateControl(); removal.PerformLayout();
            foreach (var layout in Descendants(removal).OfType<TableLayoutPanel>()) layout.PerformLayout();
            var backupToggle = Descendants(removal).OfType<CheckBox>().Single();
            backupToggle.Parent!.PerformLayout();
            Check(backupToggle.Parent.Height >= 100, "Opção de backup ficou sem espaço no diálogo.");
            Check(!backupToggle.Bounds.IntersectsWith(backupToggle.Parent.Controls.OfType<Label>().Single().Bounds), "Explicação encobriu a opção de backup MSG.");
            var input = removeControls.OfType<TextBox>().Single(x => !x.ReadOnly); input.Text = "remover"; Check(!confirm.Enabled, "Confirmação minúscula aceita.");
            input.Text = "REMOVER"; Check(confirm.Enabled, "Confirmação explícita não aceita.");
            var mode = removeControls.OfType<ComboBox>().Single(); mode.SelectedIndex = 1; Check(removal.Mode == "mailbox", "Modo de remoção da caixa não mapeado.");
            using var advanced = new RemovalDialog(new JobStatus { Id = jobs[0].Id, SourceLabel = "Caixa > Recebidos", RetainedMsgBackupForRemoval = true }, "C:\\ArquivoOutlook");
            advanced.CreateControl(); advanced.PerformLayout();
            foreach (var container in Descendants(advanced).OfType<TableLayoutPanel>()) container.PerformLayout();
            var explanation = Descendants(advanced).OfType<Label>().Single(x => x.Text.Contains("Este plano permite validação com backup MSG retido."));
            int textHeight = TextRenderer.MeasureText(explanation.Text, explanation.Font, new Size(explanation.Width, int.MaxValue), TextFormatFlags.WordBreak).Height;
            Check(textHeight <= explanation.Height, "Explicação da validação avançada cortada no diálogo de remoção.");
            ResponsiveLayout(Check);
            LargeHistory(Check);
            Console.WriteLine($"Interface: {assertions} verificações passaram. Nenhuma conexão ao Outlook ou alteração de mensagens.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine("Falha na interface: " + ex); return 1; }
    }
    private static void WizardActionTests(Action<bool,string> check)
    {
        using var form=new MainForm();form.CreateControl();
        var history=(ComboBox)typeof(MainForm).GetField("jobs",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance)!.GetValue(form)!;
        var controls=Descendants(form).OfType<Button>().ToDictionary(x=>x.Text);
        foreach(var sample in new[]{("Planned",true,false,false,false),("Copied",false,true,false,false),("ReviewRequired",false,false,true,false),("Verified",false,false,false,true),("Removed",false,false,false,false),("DeleteStarted",false,false,true,true)})
        {
            var job=new JobStatus{Id=Guid.NewGuid().ToString("N"),States=new(){{sample.Item1,5}}};history.Items.Add(job);history.SelectedItem=job;
            check(controls["Copiar / retomar"].Enabled==sample.Item2,"Disponibilidade da cópia: "+sample.Item1);
            check(controls["Verificar pendentes"].Enabled==sample.Item3,"Disponibilidade da verificação: "+sample.Item1);
            check(controls["Reconciliar pendências"].Enabled==sample.Item4,"Disponibilidade da reconciliação: "+sample.Item1);
            check(controls["Remover originais…"].Enabled==sample.Item5,"Disponibilidade da remoção: "+sample.Item1);
        }
        string folder=Path.Combine(Path.GetTempPath(),"pst-destination-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        string chosen=Path.Combine(folder,"novo.pst");check(WizardRules.ValidateNewPst(chosen)==chosen,"Novo destino local rejeitado.");File.WriteAllText(chosen,"fixture");
        try{WizardRules.ValidateNewPst(chosen);throw new Exception("PST existente aceito como novo.");}catch(ArgumentException){check(true,"Colisão rejeitada.");}
        File.Delete(chosen);Directory.Delete(folder);
    }
    private static void LargeHistory(Action<bool, string> check)
    {
        using var form = new MainForm();
        form.WindowState = FormWindowState.Normal; form.StartPosition = FormStartPosition.Manual;
        form.Location = new Point(-30000, -30000); form.Size = new Size(1240, 850); form.ShowInTaskbar = false;
        form.Show();Application.DoEvents();
        typeof(MainForm).GetMethod("ShowStep",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic)!.Invoke(form,[5]);
        var bind = typeof(MainForm).GetMethod("ShowJob", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var large = new JobStatus { Id = "0123456789abcdef0123456789abcdef", SourceLabel = "Caixa > Entrada", States = new() { ["Copied"] = 6000, ["ReviewRequired"] = 120 } };
        large.Errors.AddRange(Enumerable.Range(0,6000).Select(i => new JobError { Key = i.ToString("X64"), Error = "Cópia rápida concluída; validação completa pendente antes da remoção dos originais." }));
        large.Errors.AddRange(Enumerable.Range(6000,120).Select(i => new JobError { Key = i.ToString("X64"), Error = "Divergência que precisa de conferência. " + new string('x',300) }));
        var small = new JobStatus { Id = "1123456789abcdef0123456789abcdef", States = new() { ["Copied"] = 10 }, Errors = [new() { Key = "abc", Error = "destination_durable_reopen_or_independent_backup_missing" }] };
        var tabs = Descendants(form).OfType<TabControl>().Single();
        var watch = System.Diagnostics.Stopwatch.StartNew();
        for (int i=0;i<5;i++)
        {
            bind.Invoke(form,[large]); tabs.SelectedIndex = 2; Application.DoEvents();
            var pending = Descendants(form).OfType<DataGridView>().Single(x => x.Columns.Contains("Motivo"));
            check(pending.Rows.Count == 120, "Avisos normais de cópia lotaram a grade de pendências.");
            check(pending.GetPreferredSize(new Size(500,0)).Height <= 250, "Histórico alterou o tamanho preferido da janela para o total de linhas.");
            bind.Invoke(form,[small]); tabs.SelectedIndex = 0; Application.DoEvents();
        }
        check(watch.Elapsed < TimeSpan.FromSeconds(15), "Alternar histórico grande e abas excedeu o limite de resposta.");
        large.Errors.Clear();
        large.Errors.AddRange(Enumerable.Range(0,6000).Select(i => new JobError { Key = i.ToString("X64"), Error = "volume-reopen-pending:PST ocupado" }));
        bind.Invoke(form,[large]); tabs.SelectedIndex = 2; Application.DoEvents();
        var paged = Descendants(form).OfType<DataGridView>().Single(x => x.Columns.Contains("Motivo"));
        check(paged.Rows.Count == 200, "Milhares de erros não foram paginados.");
        Descendants(form).OfType<Button>().Single(x => x.Text == "Próxima").PerformClick();
        check(paged.Rows[0].Cells[0].Value?.ToString() == 200.ToString("X64")[..12], "Navegação de pendências não abriu a próxima página.");
        Console.WriteLine($"Histórico grande: 6.120 registros, cinco alternâncias em {watch.Elapsed.TotalMilliseconds:F0} ms.");
        // Dispose without FormClosing: test must not save user settings.
    }
    private static void NewItemsTests(Action<bool, string> check)
    {
        foreach (string command in new[] { "archive-run", "archive-verify" })
        {
            check(NewItemsFollowUp.ShouldCheck(command, new EngineResult(0, "", "")), "Conclusão não consulta novos itens.");
            check(NewItemsFollowUp.ShouldCheck(command, new EngineResult(3, "", "")), "Conclusão com pendências não consulta novos itens.");
            check(!NewItemsFollowUp.ShouldCheck(command, new EngineResult(1, "", "falha")), "Falha consulta novos itens indevidamente.");
            check(!NewItemsFollowUp.ShouldCheck(command, new EngineResult(0, "", "", true)), "Operação interrompida consulta novos itens.");
        }
        foreach (string command in new[] { "archive-plan", "archive-remove", "archive-reconcile", "archive-status" })
            check(!NewItemsFollowUp.ShouldCheck(command, new EngineResult(0, "", "")), "Comando fora do fluxo consulta novos itens.");

        const string root = "C:\\ArquivoOutlook", oldJob = "0123456789abcdef0123456789abcdef", newJob = "fedcba9876543210fedcba9876543210";
        Task.Run(async () =>
        {
            async Task<(NewItemsFollowUpResult Result, List<string[]> Calls, int Prompts)> Scenario(
                EngineResult response, bool accept = false, EngineResult? plan = null)
            {
                var calls = new List<string[]>(); int prompts = 0;
                var result = await NewItemsFollowUp.ExecuteAsync(root, oldJob, args =>
                {
                    calls.Add(args);
                    return Task.FromResult(calls.Count == 1 ? response : plan ?? new EngineResult(0, "EXECUCAO=" + newJob + "\n", ""));
                }, _ => { prompts++; return Task.FromResult(accept); });
                return (result, calls, prompts);
            }
            static EngineResult Count(int count) => new(0, JsonSerializer.Serialize(new { NewItemsCount = count, CheckedUtc = DateTimeOffset.UtcNow, Error = (string?)null }), "");
            var empty = await Scenario(Count(0), true);
            check(empty.Prompts == 0 && empty.Calls.Count == 1 && empty.Result.NewJobId is null && empty.Result.Warning is null && empty.Result.NewItemsCount == 0, "Consulta vazia pediu novo plano.");
            var declined = await Scenario(Count(3));
            check(declined.Prompts == 1 && declined.Calls.Count == 1 && declined.Result.NewItemsCount == 3 && declined.Result.NewJobId is null, "Recusa preparou ou perdeu a contagem de novos itens.");
            var accepted = await Scenario(Count(3), true);
            check(accepted.Prompts == 1 && accepted.Calls.Count == 2 && accepted.Result.NewJobId == newJob, "Aceite não preparou novo plano.");
            check(accepted.Calls[0].SequenceEqual(new[] { "archive-check-new", root, oldJob }) && accepted.Calls[1].SequenceEqual(new[] { "archive-plan-new", root, oldJob }), "Novos itens mudaram a raiz/execução de referência ou iniciaram cópia automática.");
            var warning = await Scenario(new EngineResult(0, "{\"NewItemsCount\":3,\"Error\":\"Outlook ocupado\"}", ""), true);
            check(warning.Result.Warning?.Contains("Outlook ocupado") == true && warning.Prompts == 0 && warning.Calls.Count == 1, "Erro informativo preparou novo plano.");
            var badJson = await Scenario(new EngineResult(0, "{\"outraPropriedade\":3}", ""), true);
            check(badJson.Result.Warning is not null && badJson.Prompts == 0, "Contrato inválido foi aceito como consulta vazia.");
            var failedCheck = await Scenario(new EngineResult(1, "", "origem indisponível"), true);
            check(failedCheck.Result.Warning?.Contains("origem indisponível") == true && failedCheck.Prompts == 0, "Falha de consulta propagou ou preparou plano.");
            var cancelledCheck = await Scenario(new EngineResult(-1, "", "", true), true);
            check(cancelledCheck.Result.WasInterrupted && cancelledCheck.Prompts == 0, "Consulta cancelada pediu confirmação.");
            var failedPlan = await Scenario(Count(2), true, new EngineResult(1, "", "plano indisponível"));
            check(failedPlan.Result.Warning?.Contains("plano indisponível") == true && failedPlan.Result.NewJobId is null, "Falha do novo plano foi declarada concluída.");
            var cancelledPlan = await Scenario(Count(2), true, new EngineResult(-1, "", "", true));
            check(cancelledPlan.Result.WasInterrupted && cancelledPlan.Calls.Count == 2 && cancelledPlan.Result.NewJobId is null, "Novo plano cancelado iniciou outra operação.");
            var malformedPlan = await Scenario(Count(2), true, new EngineResult(0, "EXECUCAO=inválida", ""));
            check(malformedPlan.Result.Warning is not null && malformedPlan.Result.NewJobId is null, "Identidade inválida do novo plano aceita.");
            var exception = await NewItemsFollowUp.ExecuteAsync(root, oldJob, _ => throw new IOException("Consulta indisponível"), _ => throw new InvalidOperationException("Não deveria perguntar."));
            check(exception.Warning?.Contains("Consulta indisponível") == true, "Exceção da consulta não preservou o resultado anterior.");
        }).GetAwaiter().GetResult();
    }
    private static void CancellationTests(Action<bool, string> check)
    {
        Task.Run(async () =>
        {
            var client = new EngineClient(typeof(UiTests).Assembly.Location);
            var ready = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<EngineResult>? pending = null;
            try
            {
                pending = client.RunAsync(["--self-test-worker", "wait"], line => { if (line == "READY") ready.TrySetResult(); });
                await ready.Task.WaitAsync(TimeSpan.FromSeconds(10));
                client.Interrupt(); client.Interrupt();
                var cancelled = await pending.WaitAsync(TimeSpan.FromSeconds(10));
                check(cancelled.WasInterrupted && cancelled.ExitCode != 0, "Interrupção solicitada não foi identificada.");
                check(!client.IsRunning, "Motor cancelado continua marcado em execução.");
                var failed = await client.RunAsync(["--self-test-worker", "fail"]);
                check(failed.ExitCode == -1 && !failed.WasInterrupted, "Falha real -1 foi confundida com cancelamento.");
                var success = await client.RunAsync(["--self-test-worker", "success"]);
                check(success.ExitCode == 0 && !success.WasInterrupted, "Nova operação herdou o cancelamento anterior.");
                client.Interrupt(); check(!client.IsRunning, "Interromper sem operação não foi idempotente.");
            }
            finally
            {
                client.Interrupt();
                if (pending is not null) await pending.WaitAsync(TimeSpan.FromSeconds(10));
            }
        }).GetAwaiter().GetResult();
    }
    private static void ResponsiveLayout(Action<bool,string> check)
    {
        foreach(var sample in new[]{(900,680,10f),(820,600,12f),(1100,800,15f)})
        {
            using var form=new MainForm();form.ShowInTaskbar=false;form.StartPosition=FormStartPosition.Manual;
            form.Location=new Point(-30000,-30000);form.Font=new Font("Segoe UI",sample.Item3);form.ClientSize=new Size(sample.Item1,sample.Item2);form.Show();
            Application.DoEvents();
            var show=typeof(MainForm).GetMethod("ShowStep",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic)!;
            check(!form.MaximizeBox && form.FormBorderStyle==FormBorderStyle.FixedSingle,"Janela pode ser maximizada ou redimensionada.");
            for(int step=0;step<6;step++)
            {
                show.Invoke(form,[step]);
                if(step==5)
                {
                    var sampleJob=new JobStatus{Id="ffffffffffffffffffffffffffffffff",SourceLabel="Caixa de teste > Itens Enviados",States=new(){{"Copied",5000},{"InDeletedItems",4995},{"ReviewRequired",17}},Workflow=new(){State="Paused",MoveOriginals=true,TotalItems=10012,CompleteItems=4995,ReviewItems=17,PendingItems=5000,TotalBatches=102,CompletedBatches=49,CurrentBatch=50}};
                    var jobList=(ComboBox)typeof(MainForm).GetField("jobs",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance)!.GetValue(form)!;jobList.Items.Add(sampleJob);jobList.SelectedItem=sampleJob;
                    var info=new WorkflowInfo{State="Running",TotalItems=10012,CompleteItems=4995,ReviewItems=17,PendingItems=5000,TotalBatches=102,CurrentBatch=50,ItemsPerSecond=1.2,RemainingSeconds=4167,Activity="Validando e concluindo as mensagens do lote"};
                    typeof(MainForm).GetMethod("ApplyWorkflowProgress",System.Reflection.BindingFlags.NonPublic|System.Reflection.BindingFlags.Instance)!.Invoke(form,[info,true]);
                    Descendants(form).OfType<CheckBox>().Single(x=>x.Text=="Ações avançadas").Checked=true;
                }
                for(int pass=0;pass<3;pass++){foreach(var child in Descendants(form).Reverse().Where(x=>x.HasChildren))child.PerformLayout();form.PerformLayout();Application.DoEvents();}
                foreach(var button in Descendants(form).OfType<Button>().Where(x=>x.Visible))
                {
                    var text=TextRenderer.MeasureText(button.Text,button.Font);
                    check(button.ClientSize.Width>=text.Width+button.Padding.Horizontal+6,"Botão cortado: "+button.Text);
                    check(button.ClientSize.Height>=text.Height+button.Padding.Vertical+6,"Altura cortada: "+button.Text);
                    CheckLocalContainment(button,"Etapa "+step,check);
                }
                var progress=Descendants(form).OfType<ProgressBar>().Single();
                check(progress.Width>=form.ClientSize.Width-50,"Barra de progresso não ocupa o rodapé.");
                check(!Descendants(form).OfType<CheckBox>().Any(x=>x.Text.Contains("Mostrar registro")),"Painel de log ainda aparece.");
                if(sample.Item1==900 && Environment.GetEnvironmentVariable("OUTLOOK_WIZARD_CAPTURE") is { Length:>0 } capture)
                {
                    Directory.CreateDirectory(capture);using var bitmap=new Bitmap(form.Width,form.Height);form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,form.Size));
                    bitmap.Save(Path.Combine(capture,"step-"+step+".png"));
                }
            }
            form.Hide();
        }
    }
    private static void DumpLayout(Control control, int depth)
    {
        if (!control.Visible) return;
        string table = control is TableLayoutPanel layout
            ? " rows=" + string.Join(",", layout.GetRowHeights()) + " styles=" + string.Join(",", layout.RowStyles.Cast<RowStyle>().Select(x => x.SizeType + ":" + x.Height))
            : "";
        Console.Error.WriteLine(new string(' ', depth * 2) + control.GetType().Name + " [" + Describe(control).Replace('\r', ' ').Replace('\n', ' ') + "] bounds=" + control.Bounds +
            " client=" + control.ClientSize + " preferred=" + control.GetPreferredSize(new Size(control.Width, 0)) + " min=" + control.MinimumSize + " font=" + control.Font.SizeInPoints + table);
        foreach (Control child in control.Controls) DumpLayout(child, depth + 1);
    }
    private static void ProbeNativeRows(Form form)
    {
        // Diagnostic only, after a scenario has already failed. Prove whether
        // native single-line height negotiation, rather than body height, is
        // causing an AutoSize row to retain its old font's dimensions.
        foreach (var control in Descendants(form).Where(x => x.Visible && x.Dock == DockStyle.Fill &&
            (x is ComboBox || x is TextBox { Multiline: false })).ToList())
        {
            if (control.Parent is not TableLayoutPanel table) continue;
            string before = string.Join(",", table.GetRowHeights());
            int preferredHeight = control.GetPreferredSize(Size.Empty).Height;
            control.MinimumSize = new Size(control.MinimumSize.Width, preferredHeight);
            control.Dock = DockStyle.None; control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            table.PerformLayout();
            Console.Error.WriteLine($"Native-row probe {Describe(control)} font={control.Font.SizeInPoints}: preferredHeight={preferredHeight}, rows before={before}, after={string.Join(",", table.GetRowHeights())}.");
        }
    }
    private static void CheckLocalContainment(Control control, string scenario, Action<bool, string> check)
    {
        // Scrolling a content panel outside the viewport is intentional. A
        // child escaping an ordinary non-scrolling row/container is clipping.
        for (Control child = control; child.Parent is Control parent && parent is not Form; child = parent)
        {
            if (parent is ScrollableControl { AutoScroll: true }) continue;
            check(parent.ClientRectangle.Contains(child.Bounds),
                $"{scenario}: ‘{Describe(child)}’ está cortado por ‘{Describe(parent)}’: bounds {child.Bounds}, cliente {parent.ClientRectangle}.");
        }
    }
    private static string Describe(Control control) => string.IsNullOrWhiteSpace(control.Text)
        ? control.AccessibleName ?? control.GetType().Name : control.Text;
    private static IEnumerable<Control> Descendants(Control root)
    {
        foreach (Control control in root.Controls)
        { yield return control; foreach (var child in Descendants(control)) yield return child; }
    }
}
