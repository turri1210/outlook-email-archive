# Interface por etapas

A aplicação usa uma janela fixa, sem maximização, ajustada à área útil da tela. O fluxo é: início, origem, destino, opções, resumo e execução. O usuário pode voltar às escolhas antes de preparar o plano.

- **Início:** cria a pasta `emailArchive` ao lado do executável e consulta somente esse catálogo. Permite começar ou continuar uma execução disponível. Não importa histórico de outras instalações.
- **Origem:** carrega caixas e pastas do Outlook clássico e permite filtrar e selecionar a pasta inicial.
- **Destino:** escolhe um PST existente ou a pasta e o nome de um novo arquivo. Nunca sobrescreve um PST novo com nome já ocupado. As partes adicionais permanecem na mesma pasta do PST inicial e incluem a identificação da execução no nome.
- **Opções:** subpastas, corte inclusivo por data, autorização de validação com backup MSG, tamanho alvo de partes (20 GiB) e movimento automático de originais validados para Itens Excluídos.
- **Resumo:** mostra as escolhas e oferece Preparar arquivamento. Preparar não inicia cópia nem remoção.
- **Execução:** botão principal Continuar arquivamento, pausa ao concluir o lote, progresso numérico com velocidade e estimativa, abas Mensagens, Arquivos PST e Pendências, grupos de motivos e relatórios em `emailArchive/reports`. As operações manuais ficam em Ações avançadas.

## Disponibilidade das ações

| Ação | Estados com trabalho disponível |
|---|---|
| Continuar arquivamento | Planned, Exported, ImportStarted, Copied; Verified e DeleteStarted quando o movimento está autorizado |
| Copiar / retomar (avançado) | Planned, Exported, ImportStarted |
| Verificar pendentes (avançado) | Copied; Verified e DeleteStarted quando o movimento está autorizado |
| Reconciliar pendências | ReviewRequired, ImportStarted, DeleteStarted, PurgeStarted |
| Remover originais | Verified, InDeletedItems, DeleteStarted, PurgeStarted |

Durante uma operação, as escolhas e demais ações ficam bloqueadas. Interromper fica disponível quando o motor está ativo. Acima da barra animada, que ocupa toda a largura do rodapé, aparece a atividade atual. Os registros completos de cada chamada do motor são gravados em `emailArchive/logs`, incluindo consultas e interrupções. Não existe painel de log na interface.

O movimento automático é uma opção explícita no resumo e na execução. Quando ativado, Copiar / retomar segue para a validação; cada mensagem aprovada é movida para Itens Excluídos nessa mesma passagem. Verificar pendentes também trata originais já validados e ainda não movidos, após conferência atual. A gravação da intenção e da prova precede o movimento; a retomada reconhece mensagens já movidas. Não há exclusão definitiva nesse modo. A remoção definitiva permanece manual. Uma cópia sem validação não habilita remoção. Rascunhos, tipos não suportados e itens com divergências não são aprovados pelo redesenho da interface.

## Distribuição

Copie a pasta publicada inteira. Ela contém o executável da interface, suas dependências e a subpasta `engine`, com o motor, os três auxiliares nativos e os recursos de validação. O executável encontra o motor nessa subpasta sem depender do diretório do código-fonte. A instalação atual usa .NET 8 e o runtime Python configurado neste computador; as dependências Python da validação estão incluídas em `engine/semantic`.

## Fluxo por lotes

### Retomar a preparação

Ao iniciar um planejamento, a execução é registrada antes da leitura dos e-mails.
As identidades de cada pasta capturada e o resultado da análise de cada mensagem
são salvos no catálogo. Ao interromper e reabrir, escolha **Continuar backup** e
depois **Continuar arquivamento** para concluir a preparação da mesma execução.
Os itens já analisados não são relidos, e novas mensagens não entram nas pastas
cujo escopo já foi capturado. A captura de uma pasta interrompida antes de seu
checkpoint precisa ser refeita; as pastas já salvas são reaproveitadas.

Até a conclusão do plano, cópia, validação, reconciliação e remoção permanecem
bloqueadas. O término da preparação não inicia a cópia automaticamente. Uma pasta
ou mensagem do escopo que tenha sido movida ou excluída pode exigir revisão;
a retomada não a substitui por outra mensagem.

O comando técnico de retomada é `archive-plan-resume <raiz> <execução>`.
`archive-planning-self-test` exercita interrupções na captura, na análise e antes
do registro final, usando apenas mensagens artificiais.

O plano inicial permanece congelado. O manifesto `jobs/<execução>/workflow/manifest.json` divide os itens em grupos de até 100, preservando inteiros os lotes nativos que já existiam. O banco registra os membros e o checkpoint de cada lote em transações duráveis.

Continuar arquivamento confere a conexão com o Outlook clássico, a presença dos PSTs, o acesso às pastas e a reserva de espaço. Depois processa somente os lotes abertos: copiar o que falta, validar as cópias e mover apenas os aprovados, quando autorizado. Lotes concluídos não têm seu conteúdo reaberto; consultas a índices e contagens continuam necessárias. Pendências ficam em lotes para revisão e não impedem o avanço nos demais.

O botão também fica disponível quando só há itens para reconciliar ou remover. Ao acioná-lo, cada lote com pendências é reconciliado uma vez; os itens recuperados seguem para validação e movimento autorizado no mesmo fluxo. Divergências que persistirem continuam em Pendências, sem repetição automática. Os estados atuais das mensagens prevalecem sobre um checkpoint antigo do lote. Quando restar apenas remoção manual, Continuar abre a confirmação de remoção existente; a exclusão definitiva nunca é iniciada automaticamente.

Pausar após este lote cria uma solicitação local que o motor reconhece depois de gravar o checkpoint do lote. Interromper continua disponível para parar imediatamente; nesse caso a retomada usa as intenções por mensagem do lote interrompido. Falhas que deixam itens sem processamento pausam o fluxo, sem repetição em ciclo. O usuário retoma após conferir o Outlook.

A versão do catálogo passa a exigir o motor com suporte ao fluxo por lotes. Um backup anterior à conversão é preservado para diagnóstico; ele não deve substituir o catálogo após novas movimentações, pois isso descartaria checkpoints posteriores.

Na conversão, erros específicos de identificação nativa temporariamente indisponível podem voltar de revisão para Copiado quando as evidências completas já estão registradas. Isso não aprova nenhuma mensagem: uma nova validação completa continua obrigatória antes de qualquer movimento. Outras divergências permanecem em revisão.

Ao terminar, o programa apresenta e salva o resumo das cópias, movimentações e pendências. A estimativa de duração é calculada com a velocidade da execução atual e pode variar conforme o tamanho e a complexidade das mensagens.
