namespace OutlookArchive.UI;

internal sealed partial class MainForm
{
    private readonly Panel pagesHost=new(){Dock=DockStyle.Fill};
    private readonly List<Panel> pages=[];
    private readonly Label stepHeading=new(){AutoSize=true,Font=new Font("Segoe UI",18,FontStyle.Bold)};
    private readonly Label stepDescription=new(){AutoSize=true,ForeColor=Color.FromArgb(75,89,104)};
    private readonly Label review=new(){AutoSize=true,Dock=DockStyle.Top};
    private readonly Label resumeLabel=WrappedLabel("Execuções registradas nesta instalação:");
    private readonly CheckBox autoMoveNew=new(){Text="Mover originais para Itens Excluídos após validar",AutoSize=true,Checked=true};
    private readonly CheckBox autoMoveCurrent=new(){Text="Mover originais para Itens Excluídos após validar",AutoSize=true};
    private bool loadingAutoMovement;
    private readonly Button back=Button("Voltar"), next=Button("Avançar",true), home=Button("Voltar ao início");
    private readonly Button startNew=Button("Iniciar um backup",true), resume=Button("Continuar backup");
    private readonly Button selectNewPst=Button("Escolher pasta e nome…");
    private readonly TextBox newPst=new(){Dock=DockStyle.Fill,ReadOnly=true,PlaceholderText="Escolha onde criar o arquivo PST",AccessibleName="Caminho do novo PST"};
    private int currentStep;
    private readonly string[] titles=["Como deseja começar?","1. Escolha a origem","2. Escolha o destino","3. Opções do backup","4. Confira o arquivamento","5. Execução e validação"];
    private readonly string[] descriptions=["Inicie um backup ou continue uma execução registrada nesta instalação.","Selecione a pasta do Outlook clássico a partir da qual deseja arquivar.","Crie um PST ou escolha um arquivo já existente. O conteúdo existente será preservado.","Defina o conteúdo e os limites do arquivamento.","Confirme as opções abaixo. Preparar o plano ainda não inicia a cópia.","Acompanhe as mensagens e trate apenas as etapas que possuem itens disponíveis."];

    private void BuildWizard()
    {
        var shell=new TableLayoutPanel{Dock=DockStyle.Fill,Padding=new Padding(20),ColumnCount=1,RowCount=5};
        shell.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
        shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));shell.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));shell.RowStyles.Add(new RowStyle(SizeType.AutoSize));shell.RowStyles.Add(new RowStyle(SizeType.Absolute,8));
        var heading=NaturalTable(1,2);heading.Dock=DockStyle.Top;heading.Controls.Add(stepHeading,0,0);heading.Controls.Add(stepDescription,0,1);BoundLabel(stepDescription);
        shell.Controls.Add(heading,0,0);shell.Controls.Add(pagesHost,0,1);
        for(int i=0;i<6;i++){var page=new Panel{Dock=DockStyle.Fill,Visible=false,Padding=new Padding(0,16,0,8)};pages.Add(page);pagesHost.Controls.Add(page);}

        var welcome=NaturalTable(1,6);welcome.Dock=DockStyle.Top;
        welcome.Controls.Add(WrappedLabel("O Outlook clássico deve estar aberto. Você pode mover os originais para Itens Excluídos automaticamente após validar as cópias."),0,0);
        welcome.Controls.Add(startNew,0,1);welcome.Controls.Add(resumeLabel,0,2);
        welcome.Controls.Add(jobs,0,3);welcome.Controls.Add(resume,0,4);
        welcome.Controls.Add(WrappedLabel("Os arquivos de trabalho e os logs ficam na pasta emailArchive, ao lado do programa."),0,5);pages[0].Controls.Add(welcome);
        var sourcePage=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=3};sourcePage.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
        sourcePage.RowStyles.Add(new RowStyle(SizeType.AutoSize));sourcePage.RowStyles.Add(new RowStyle(SizeType.AutoSize));sourcePage.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        sourcePage.Controls.Add(connect,0,0);sourcePage.Controls.Add(search,0,1);sourcePage.Controls.Add(sources,0,2);pages[1].Controls.Add(sourcePage);
        destinationMode.Items.AddRange(["Criar um novo PST","Usar um PST existente"]);destinationMode.SelectedIndex=0;
        var target=NaturalTable(1,7);target.Dock=DockStyle.Top;
        target.Controls.Add(destinationMode,0,0);target.Controls.Add(WrappedLabel("Novo arquivo PST"),0,1);target.Controls.Add(newPst,0,2);target.Controls.Add(selectNewPst,0,3);
        target.Controls.Add(WrappedLabel("Arquivo PST existente"),0,4);target.Controls.Add(existingPst,0,5);target.Controls.Add(browsePst,0,6);pages[2].Controls.Add(target);
        var options=NaturalTable(1,7);options.Dock=DockStyle.Top;
        options.Controls.Add(autoMoveNew,0,6);
        options.Controls.Add(includeChildren,0,0);options.Controls.Add(cutoffEnabled,0,1);options.Controls.Add(cutoffPicker,0,2);options.Controls.Add(retainedMsgBackup,0,3);
        options.Controls.Add(WrappedLabel(UiRules.RetainedMsgExplanation),0,4);
        var size=new FlowLayoutPanel{AutoSize=true,Dock=DockStyle.Top};size.Controls.Add(new Label{Text="Dividir PSTs a cada",AutoSize=true,Margin=new Padding(0,6,8,0)});size.Controls.Add(volumeSize);size.Controls.Add(new Label{Text="GiB",AutoSize=true,Margin=new Padding(8,6,0,0)});options.Controls.Add(size,0,5);pages[3].Controls.Add(options);
        pages[4].AutoScroll=true;pages[4].Controls.Add(review);BoundLabel(review);
        var execution=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=5};execution.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
        execution.RowStyles.Add(new RowStyle(SizeType.AutoSize));execution.RowStyles.Add(new RowStyle(SizeType.Percent,100));execution.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        execution.RowStyles.Add(new RowStyle(SizeType.AutoSize));execution.Controls.Add(autoMoveCurrent,0,2);
        execution.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        autoMoveCurrent.CheckedChanged+=(_,_)=>{
            if(!loadingAutoMovement && !busy && displayedJob is { } job)
            {
                try{RemovalPreferences.Save(root.Text,job.Id,autoMoveCurrent.Checked,Path.Combine(root.Text,"preferences","automatic-movement"));}
                catch(IOException ex){AppendLog("Não foi possível lembrar o movimento automático: "+ex.Message);}
                catch(UnauthorizedAccessException ex){AppendLog("Não foi possível lembrar o movimento automático: "+ex.Message);}
            }
            UpdateActions();
        };
        summary.Dock=DockStyle.Top;summary.MaximumSize=Size.Empty;execution.Controls.Add(summary,0,0);BoundLabel(summary);
        var tabs=new TabControl{Dock=DockStyle.Fill};var states=new TabPage("Mensagens");states.Controls.Add(stateGrid);var volumes=new TabPage("Arquivos PST");volumes.Controls.Add(volumeGrid);
        var pending=new TabPage("Pendências");var pendingPanel=new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=3,ColumnCount=1};pendingPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));pendingPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));pendingPanel.RowStyles.Add(new RowStyle(SizeType.Percent,100));pendingPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        pendingPanel.Controls.Add(pendingGroups,0,0);pendingGroups.SelectedIndexChanged+=(_,_)=>{if(!loadingPendingGroups){pendingPage=0;ShowPendingPage();}};
        pendingPanel.Controls.Add(errorGrid,0,1);var pendingTools=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true,WrapContents=true};pendingTools.Controls.AddRange([previousPending,pendingPageLabel,nextPending,pendingHelp,pendingReport]);BoundFlow(pendingTools);pendingPanel.Controls.Add(pendingTools,0,2);pending.Controls.Add(pendingPanel);
        tabs.TabPages.AddRange([states,volumes,pending]);execution.Controls.Add(tabs,0,1);
        var actions=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true,WrapContents=true,Padding=new Padding(0,8,0,0)};actions.Controls.AddRange([continueArchive,pauseAfterBatch,advancedActions]);BoundFlow(actions);execution.Controls.Add(actions,0,3);
        var advanced=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true,WrapContents=true,Visible=false};advanced.Controls.AddRange([copy,verify,reconcile,remove,retainedMsgCurrent]);execution.Controls.Add(advanced,0,4);BoundFlow(advanced);
        hints.SetToolTip(retainedMsgCurrent,UiRules.RetainedMsgExplanation);
        retainedMsgCurrent.CheckedChanged+=(_,_)=>{
            if(!loadingAutoMovement && !busy && displayedJob is { } current)
            {
                try{RemovalPreferences.Save(root.Text,current.Id,retainedMsgCurrent.Checked);}
                catch(Exception ex) when(ex is IOException or UnauthorizedAccessException){AppendLog("Não foi possível salvar a autorização MSG: "+ex.Message);}
            }
        };
        advancedActions.CheckedChanged+=(_,_)=>advanced.Visible=advancedActions.Checked;
        continueArchive.Click+=async (_,_)=>await ContinueArchive();pauseAfterBatch.Click+=(_,_)=>RequestBatchPause();pages[5].Controls.Add(execution);
        var navigation=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true,WrapContents=true,FlowDirection=FlowDirection.RightToLeft,Padding=new Padding(0,8,0,0)};
        navigation.Controls.AddRange([next,plan,back,home,stop]);BoundFlow(navigation);shell.Controls.Add(navigation,0,2);
        activity.AutoSize=true;activity.Dock=DockStyle.Top;activity.Margin=new Padding(0,8,0,8);BoundLabel(activity);shell.Controls.Add(activity,0,3);
        progress.Dock=DockStyle.Fill;progress.Margin=Padding.Empty;progress.MarqueeAnimationSpeed=30;shell.Controls.Add(progress,0,4);Controls.Add(shell);
        startNew.Click+=async (_,_)=>{ShowStep(1);if(sourceData.Count==0)await Guarded("Carregando pastas do Outlook…",LoadSources);};
        resume.Click+=(_,_)=>{if(jobs.SelectedItem is JobStatus){ShowSelectedJob();ShowStep(5);activity.Text="Use Continuar arquivamento para retomar os lotes pendentes.";}};
        home.Click+=async (_,_)=>await Guarded("Consultando execuções…",async()=>{await RefreshJobs();ShowStep(0);});
        back.Click+=(_,_)=>ShowStep(Math.Max(0,currentStep-1));
        next.Click+=(_,_)=>{
            try{if(currentStep==2)ValidateDestination();if(currentStep==3)BuildReview();ShowStep(currentStep+1);}
            catch(Exception ex){MessageBox.Show(this,ex.Message,"Confira as opções",MessageBoxButtons.OK,MessageBoxIcon.Information);}
        };
        selectNewPst.Click+=(_,_)=>{
            using var dialog=new SaveFileDialog{Title="Escolha a pasta e o nome do novo PST",Filter="Arquivo de dados do Outlook (*.pst)|*.pst",DefaultExt="pst",AddExtension=true,OverwritePrompt=false,FileName="backup-"+DateTime.Today.ToString("yyyy-MM-dd")+".pst"};
            if(dialog.ShowDialog(this)==DialogResult.OK){newPst.Text=dialog.FileName;UpdateActions();}
        };
        existingPst.TextChanged+=(_,_)=>UpdateActions();
        ShowStep(0);
    }
    private string ValidateDestination()=>destinationMode.SelectedIndex==1 ? UiRules.ValidateExistingPstPath(true,existingPst.Text)! : WizardRules.ValidateNewPst(newPst.Text);
    private void BuildReview()
    {
        string destination=ValidateDestination();
        review.Text=$"Origem\n{(sources.SelectedNode?.Tag as SourceSelection)?.Label}\n\nDestino\n{destination}\n{(destinationMode.SelectedIndex==1?"Usar arquivo existente; conteúdo preservado.":"Criar um novo arquivo PST.")}\nAs partes adicionais serão criadas na mesma pasta.\n\nOpções\nSubpastas: {(includeChildren.Checked?"incluídas":"não incluídas")}\n{UiRules.CutoffLabel(cutoffEnabled.Checked?DateOnly.FromDateTime(cutoffPicker.Value):null,TimeZoneInfo.Local.Id)}\nValidação com backup MSG: {(retainedMsgBackup.Checked?"permitida":"desativada")}\nDividir PSTs a cada {volumeSize.Value} GiB\nMovimento após validar: {(autoMoveNew.Checked?"automático para Itens Excluídos":"originais permanecem na pasta de origem")}\n\nArquivos de trabalho\n{root.Text}\n\nPreparar registra o plano. Copiar / retomar iniciará a cópia e, se autorizado acima, a validação com movimento automático.";
    }
    private void ShowStep(int step)
    {
        currentStep=Math.Clamp(step,0,5);
        for(int i=0;i<pages.Count;i++)pages[i].Visible=i==currentStep;
        pages[currentStep].BringToFront();stepHeading.Text=titles[currentStep];stepDescription.Text=descriptions[currentStep];UpdateActions();
    }
    private void UpdateWizardActions()
    {
        if(pages.Count==0)return;
        startNew.Enabled=!busy;resume.Enabled=!busy && jobs.SelectedItem is JobStatus selected && (selected.PlanningPending||WizardRules.CanContinue(selected.States));
        bool hasPending=jobs.Items.Cast<JobStatus>().Any(x=>x.PlanningPending||WizardRules.CanContinue(x.States));
        resume.Visible=jobs.Visible=hasPending;
        resumeLabel.Text=hasPending ? "Execuções registradas nesta instalação:" : "Nenhum backup pendente nesta instalação.";
        back.Visible=currentStep is >=1 and <=4;back.Enabled=!busy;
        next.Visible=currentStep is >=1 and <=3;next.Enabled=!busy && (currentStep!=1 || sources.SelectedNode?.Tag is SourceSelection);
        plan.Visible=currentStep==4;home.Visible=currentStep==5;home.Enabled=!busy;
        stop.Visible=currentStep==5||busy;
        newPst.Enabled=selectNewPst.Enabled=!busy && destinationMode.SelectedIndex==0;
    }
}
