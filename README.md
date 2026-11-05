# Arquivamento do Outlook para PST

Aplicativo local para Windows que usa o perfil já aberto no **Outlook clássico**. A interface WinForms roda em uma janela separada; não é um suplemento. O novo Outlook não fornece a conexão utilizada pela ferramenta.

Ao mover ou reinstalar a ferramenta, inclua os três componentes `native-copy.exe`, `native-identity.exe` e `native-finalize.exe`. Os artefatos devem estar em `src/native/win-x64` para compilar e junto de `OutlookArchive.dll` para executar. Eles não são versionados no Git; podem ser gerados com `tools/native/build-runtime.ps1 -ToolchainRoot <diretório do compilador>`. A compilação e a interface bloqueiam operações quando esses componentes estão ausentes. O iniciador recompila também quando faltam componentes na pasta de execução.

## O que está implementado

- Seleção de caixa/pasta e subpastas, plano de origem congelado e histórico por execução.
- Captura inicial por identificadores: novas chegadas não interrompem o planejamento. Ao terminar a cópia/verificação, a interface oferece preparar um plano complementar apenas para os novos e-mails elegíveis, sem repetir os anteriores.
- Cópia nativa em lotes de até 100 mensagens, com limite de 128 MiB por lote rápido (lotes antigos mantêm o protocolo de 50/64 MiB), para mensagens elegíveis. O fluxo rápido não exporta MSG nem compara corpos durante a transferência. A identificação provisória usa metadados; a verificação detalhada, com backup MSG, é uma etapa explícita anterior à remoção. Backups antigos continuam preservados e o transporte individual permanece para os demais casos. Mantém a hierarquia das pastas e aceita PST novo ou existente.
- Catálogo SQLite com retomada após interrupções e prevenção de duplicatas pela identidade da origem no mesmo catálogo.
- Divisão em PSTs com alvo padrão de **20 GiB**, limite interno de **45 GiB** e reserva de **512 MiB**. A estimativa conservadora pode produzir partes menores.
- Validação de datas/horários, remetente, destinatários, CC/CCO, anexos e informações do corpo antes de habilitar a remoção. Diferenças internas de HTML/RTF podem ser aceitas quando texto, links, imagens e estrutura relevante permanecem iguais; formatos sem interpretação segura continuam sujeitos à comparação exata.
- Movimento automático opcional para Itens Excluídos imediatamente após validação completa, com prova durável e retomada. A remoção definitiva dos itens rastreados dessa pasta continua sendo uma etapa manual. Nunca esvazia a pasta inteira.
- Interface para preparar, copiar/retomar, verificar, reconciliar pendências e acompanhar arquivos e estados.

**Rascunhos permanecem pendentes**, conforme o escopo combinado. Tipos diferentes de mensagens comuns `IPM.Note` também são registrados como pendentes e preservados. A implementação não representa homologação de caixas completas de 10–50 GB. Nenhum original real da empresa foi removido nos testes de desenvolvimento; os testes de remoção usam mensagens artificiais em PST local.

## Abrir e usar

1. Abra o Outlook clássico no perfil desejado.
2. Na pasta da versão compilada, abra **OutlookArchive.UI.exe**. Mantenha os demais arquivos e a pasta `engine` juntos.
3. Escolha **Iniciar um backup** ou **Continuar backup**, quando houver uma execução pendente nesta instalação.
4. Selecione a origem, o PST novo ou existente, as opções e confira o resumo.
5. Use **Preparar arquivamento**. Depois acompanhe as abas **Mensagens**, **Arquivos PST** e **Pendências**.
6. Use **Continuar arquivamento** para executar cópia, validação e movimento autorizado em lotes de até 100 mensagens. Lotes concluídos são pulados na retomada; pendências são separadas por motivo. **Pausar após este lote** conclui o lote atual antes de parar. As operações manuais ficam em **Ações avançadas**. Os botões só ficam disponíveis quando há itens elegíveis. Com **Mover originais para Itens Excluídos após validar** ativo, a cópia segue para validação e cada original aprovado é movido. A exclusão definitiva continua exigindo confirmação separada.

A janela é fixa e não pode ser maximizada. A pasta `emailArchive`, ao lado do executável, é criada automaticamente e concentra catálogo, planos, evidências, preferências, relatórios e logs. O PST e suas partes adicionais ficam na pasta escolhida no assistente. Não há painel de registro completo na interface: consulte `emailArchive/logs`.


O destino contém PSTs, catálogo e backups MSG com conteúdo dos e-mails; mantenha a pasta completa em armazenamento autorizado. O arquivo PST pode ser fechado no Outlook sem apagá-lo do disco. Evite editar os PSTs da execução em paralelo com a ferramenta.

Requisitos desta distribuição: Windows, Outlook clássico e .NET 8/Desktop Runtime 8. O iniciador compila os binários quando ausentes; essa compilação exige SDK .NET 8 ou posterior, já instalado neste computador. Execute a ferramenta com o mesmo usuário e nível de permissão do Outlook. Não é solicitada senha de e-mail.

Guia da interface: [docs/INTERFACE.md](docs/INTERFACE.md). Funcionamento, comandos e limites: [docs/MOTOR-E-VALIDACAO.md](docs/MOTOR-E-VALIDACAO.md).

## Limites importantes

A conferência compara os campos capturados e aplica equivalências de representação somente nos formatos suportados. Ela não certifica todas as propriedades MAPI, renderização visual, completude no servidor, itens ocultos/recuperáveis ou liberação de quota. A remoção em Exchange/Microsoft 365 exige origem on-line sem cache nesta versão. Mover para Itens Excluídos não libera a quota, e retenções do Microsoft 365 podem continuar aplicáveis.

Por padrão, a prova de persistência depende da reabertura e leitura exclusiva do PST exigidas pelo motor. Se o Outlook mantiver o arquivo ocupado, a validação pode continuar pendente. A alternativa por backup MSG retido está desligada por padrão e pode ser habilitada explicitamente ao preparar o plano. Ela exige o backup íntegro e o PST reaberto/revalidado; foi exercitada nos testes artificiais. Dois arquivos no mesmo disco não oferecem redundância contra falha física desse disco.

Diferenças de representação de corpo são conferidas pelo perfil de informações relevantes quando suportado. Um diagnóstico antigo não autoriza exclusão: a origem atual, o backup íntegro e o destino reaberto precisam passar na validação atual.

### Cópia em lotes e interrupções

Use **Copiar / retomar** na mesma execução; não é necessário preparar outro plano. O programa registra o escopo e a intenção antes de copiar, confere a quantidade e a identificação provisória das cópias, identifica cada uma e finaliza as pastas dentro do mesmo PST. Mensagens de conteúdo idêntico continuam sendo mensagens distintas. Um resultado parcial ou ambíguo fica em revisão e nunca é reenviado às cegas. **Reconciliar pendências** pode reconhecer correspondências de um lote parcial, mantendo somente as mensagens divergentes em revisão. Os demais lotes continuam. O comando de cópia não executa a verificação detalhada automaticamente.

O componente nativo pode terminar a chamada que já estava em andamento depois de uma interrupção. Enquanto isso, a retomada aguarda a liberação do lote e a validação não fecha o PST. A remoção dos originais continua separada e exige as mesmas provas de integridade e persistência.

Na primeira utilização de lotes, o catálogo passa a exigir esta versão ou mais recente. Os componentes nativos x64 de cópia, identificação e finalização interna do PST e os recursos em `semantic` acompanham o motor. A publicação dos recursos locais é feita por `tools/Publish-SemanticValidation.ps1`; o runtime Python configurado é necessário para a verificação detalhada do corpo. A leitura nativa em lote consulta metadados sem modificar os originais; a finalização só movimenta cópias dentro do mesmo PST.

## Desenvolvimento e testes

```powershell
dotnet build .\src\OutlookArchive.csproj -c Release
dotnet build .\ui\OutlookArchive.UI.csproj -c Release
dotnet .\src\bin\Release\net8.0-windows\OutlookArchive.dll archive-self-test
dotnet .\ui\bin\Release\net8.0-windows\OutlookArchive.UI.dll --self-test
```

Os testes acima usam dados artificiais/regras locais. Os comandos `archive-ui-sources`, `archive-plan`, `archive-run`, `archive-verify`, `archive-reconcile` e `archive-remove` conectam ao Outlook real. `archive-status` consulta o catálogo sem conectar ao Outlook.

O projeto mantém comandos `pilot-*`, inventário e diagnósticos históricos para reproduzir a investigação. Eles não substituem o motor `archive-*`. Seus relatórios antigos permanecem imutáveis e podem descrever limitações anteriores à implementação atual.

## Continuar em um arquivo existente

No campo **Destino das cópias**, escolha **Usar um PST existente** e selecione o arquivo. O histórico fica na pasta de dados criada automaticamente. O conteúdo anterior permanece preservado; novas execuções mantêm a estrutura da origem, usando as pastas padrão no idioma do PST. Quando necessário, novas partes são criadas na pasta escolhida para o PST. Não selecione o próprio armazenamento de origem nem o arquivo de dados padrão do perfil. PST ANSI, protegido, em OneDrive ou redirecionado não é aceito nesta opção.

A deduplicação acompanha as identidades transferidas pela ferramenta no mesmo catálogo. Cópias manuais preexistentes não são automaticamente adotadas nem usadas para autorizar exclusão; podem existir mensagens equivalentes em pastas distintas.

## Armazenamento e retomada otimizados

A busca de originais por identidade usa um índice de metadados por pasta, lido em blocos. Resultados negativos são atualizados no máximo uma vez por pasta e lote; correspondências ambíguas continuam bloqueadas. Novas mensagens não entram no plano atual.

Durante a conferência de uma mensagem, o MSG original permanece protegido contra escrita e exclusão. Sua leitura e impressão digital podem ser reutilizadas somente nesse intervalo; a origem atual e o destino continuam sujeitos à validação. A retomada nunca usa uma aprovação temporária de outra execução.

Ao encerrar cada lote, as evidências dos itens concluídos recebem compressão transparente do NTFS quando disponível. Caminhos, bytes lógicos e provas de integridade são preservados. Apenas cópias temporárias de trabalho identificadas e liberadas pelo Outlook são descartadas. Evidências e arquivos temporários dos itens pendentes permanecem preservados; PSTs não são comprimidos nem removidos por essa manutenção.

Eventos antigos do catálogo são armazenados em blocos GZip com verificação SHA-256, mantendo conteúdo, ordem e consulta do histórico. Os 2.000 eventos recentes permanecem diretamente na tabela. O catálogo compactado exige esta versão ou mais recente (formato 5). A manutenção de recuperação de espaço executa VACUUM somente com a automação parada.

Com o aplicativo fechado, a manutenção pode ser executada por `archive-maintain-storage <raiz> <execução>`. O comando `archive-storage-self-test` verifica índice, proteção das evidências, compactação, recuperação do histórico e preservação de arquivos pendentes usando dados artificiais. `archive-source-index-audit <raiz> <execução>` consulta até 100 originais pendentes sem alterar mensagens ou estados.
