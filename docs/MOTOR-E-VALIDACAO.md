# Motor, validação e retomada

## Conferência imediata no transporte individual

Quando o movimento automático e a prova adicional MSG retida estão autorizados,
o fluxo reutiliza os snapshots produzidos na exportação e na leitura da cópia
gravada no PST. A leitura do destino reabre o item persistido pelo identificador
e confere seu token, armazenamento e pasta. O MSG retido permanece protegido contra
escrita/exclusão enquanto seu SHA-256 é calculado em paralelo com a consulta do
ChangeKey da origem. Apenas arquivos locais passam pela thread de trabalho.

O novo atestado de leitura do item armazenado, combinado com MSG independente
íntegro, é válido somente para mover a Itens Excluídos. Não autoriza exclusão
definitiva. A validação de conteúdo, destinatários, datas e anexos permanece
obrigatória. O backend ainda confere o ChangeKey e o conteúdo após gravar o
marcador de recuperação e antes de mover. Aprovação e intenção são persistidas
antes da movimentação. Uma interrupção descarta os atalhos em memória: a retomada
usa as evidências duráveis e a conferência integral existente.

`archive-immediate-self-test` verifica fluxo normal, ChangeKey ausente, retomada
após movimento, origem/destino alterados, MSG adulterado e ausência de autorização,
com mensagens artificiais. O transporte nativo em lote mantém suas regras próprias.

## Diagnóstico detalhado de desempenho

`archive-performance-audit <raiz> <execução> <nova-pasta-local>` exige a operação
parada e o bloqueio exclusivo do catálogo. O comando executa e registra **três
cópias reais pendentes do plano existente**, sem mover os originais; depois faz
leituras de uma amostra de até 12 divergências. Não é um comando inteiramente
somente leitura. Ele não aprova divergências nem altera as regras de remoção.

`timings.jsonl` registra tempos hierárquicos da exportação MSG, metadados/anexos,
perfil do corpo, importação no PST, leitura de destino e catálogo. Os totais já
incluem seus componentes, portanto não se devem somar ambos. A instrumentação
fica desativada fora da sessão de diagnóstico. A pasta de diagnóstico pode
conter cópias temporárias de e-mails e deve permanecer em armazenamento privado.

`archive-performance-self-test` valida a hierarquia das medições e o limite de
três cópias com originais preservados, usando somente dados artificiais.
`tools/audit_archive_divergences.py <raiz> <execução> <relatório.json>` consulta
o catálogo em modo somente leitura e compara os nomes dos campos divergentes
nos perfis já salvos, sem aprovar ou modificar mensagens.

## Separação entre interface e Outlook

A interface WinForms inicia processos do motor C#/.NET com argumentos separados, sem montar comandos de shell a partir dos campos do usuário. O motor se conecta à sessão COM já aberta do Outlook clássico, no mesmo usuário e nível de permissão. Ele não inicia outro perfil, não usa o novo Outlook e não solicita senha de e-mail. As operações sobre Outlook/PST são sequenciais em STA.

O iniciador **Abrir Arquivo de Emails.cmd**, na raiz do projeto, abre a interface. O guia de operação está em [INTERFACE.md](INTERFACE.md).

## Catálogo e plano

Cada pasta de trabalho contém `archive.sqlite`, `archive.lock` e `jobs/<id>`. SQLite usa WAL, chaves estrangeiras e gravação síncrona completa; a alteração de estado e seu evento pertencem à mesma transação. Um bloqueio exclusivo por raiz impede dois executores da ferramenta sobre o mesmo catálogo, mantendo a consulta de status disponível.

O planejamento lê a origem acessível, exclui pastas virtuais de pesquisa e congela os descritores em `source-plan.json`, cujo SHA-256 fica registrado com a configuração. Retomar usa esse plano, sem redescobrir a caixa e incorporar itens novos silenciosamente. O limite padrão de descoberta é de um milhão de itens; ao excedê-lo o plano parcial não é executado.

A identidade global do catálogo é SHA-256 da composição delimitada de StoreID e EntryID da origem. Um novo plano na mesma raiz reconhece origens já registradas em outra execução. A prevenção de duplicatas não procura cópias antigas em PSTs externos nem funciona entre catálogos independentes. Mensagens de conteúdo igual, mas identidades diferentes, são preservadas separadamente.

## Cópia e recuperação

O motor exporta cada mensagem suportada para MSG Unicode e guarda seu hash e os hashes dos campos observados. Antes de importar, confere novamente a origem. A intenção de importação e a parte de destino são gravadas antes da chamada ao Outlook. A cópia armazenada recebe um token único persistido e o PST recebe identificação de posse da execução.

Se o processo for interrompido, a retomada busca o token antes de importar novamente. Uma cópia sem token só pode ser reconciliada no destino próprio, por candidato único, íntegro, não associado a outra origem, com comparação completa dos campos observados. Dois candidatos possíveis ou tokens repetidos bloqueiam a operação; não há deduplicação por assunto/data ou exclusão automática de cópias ambíguas.

O Outlook pode manter arquivos MSG ocupados mesmo após liberar o objeto COM. A conferência do backup usa cópias temporárias separadas para leitura estruturada e abertura pelo Outlook, preservando o MSG retido. Arquivos temporários bloqueados podem permanecer como evidência; não se deve apagar a pasta da execução para resolver uma falha.

## Partes PST

| Parâmetro padrão | Valor |
| --- | --- |
| Alvo de cada parte | 20 GiB |
| Limite interno configurado | 45 GiB |
| Reserva do limite | 512 MiB |
| Reserva estimada por item | Maior entre 1 MiB e 125% do tamanho observado |

O motor considera o maior valor entre tamanho físico observado e base inicial mais reservas acumuladas. Isso evita depender do crescimento físico imediato do PST, que pode ser adiado pelo Outlook. A abordagem é conservadora e pode criar arquivos menores que o alvo. Um item acima do alvo pode receber parte dedicada se ainda couber no limite utilizável; se não couber, é bloqueado. Estimativas não garantem o tamanho final do PST.

Após importação e reabertura, o tamanho observado é conferido novamente. Uma parte que viola o limite vai para revisão e não autoriza remoção. Nomes das partes usam a identidade da execução e número sequencial; PSTs alheios não são adotados como destinos.

## O que a validação comprova

A comparação é exata, sem normalizar texto, ignorar datas ou aceitar apenas igualdade de contagem. Ela compara a origem atual com a evidência inicial, o backup MSG com seu hash registrado, a cópia atual com sua evidência anterior e a origem com o destino.

Os campos capturados incluem classe, assunto, remetente/destinatários textuais, datas, leitura, importância, sensibilidade, quantidade de anexos, nomes, bytes/SHA-256 e metadados de associação de anexos. Para o corpo, são comparadas presença, tamanho e SHA-256 dos streams nativos de texto, HTML e RTF serializados no MSG. A ausência de um stream é explícita e a ausência de todos os corpos reconhecidos impede a conferência. Nenhum HTML é renderizado e nenhum anexo é executado.

Além da igualdade dos campos, a validação exige identidades conhecidas e distintas, cópia única, PST local pertencente à execução, origem na pasta esperada, mensagem `IPM.Note` enviada/recebida, backup íntegro e fase durável compatível.

Por padrão, a persistência exige desanexar/reabrir o PST e obter leitura exclusiva do arquivo conforme a prova exigida pelo motor. Se o Outlook não liberar o arquivo, a cópia pode ficar com validação pendente. A configuração técnica `AllowRetainedMsgBackupForRemoval` permite uma alternativa com backup MSG íntegro e destino reaberto/revalidado; seu padrão é `false` e a interface permite escolhê-la explicitamente no preparo do plano. Os testes a utilizaram no laboratório. A escolha fica no plano congelado e aparece na confirmação de remoção. Essa alternativa não oferece redundância física quando os arquivos estão no mesmo disco.

Essas provas se referem aos campos capturados. Não certificam todas as propriedades MAPI, fidelidade da serialização inicial ao servidor, aparência renderizada, completude da caixa no Exchange, regras, permissões, pastas ocultas/recuperáveis ou quota liberada. Diferenças observadas no Junk Email continuam sujeitas a revisão estrita.

## Remoção e interrupções

A remoção exige comando/modo explícito e novas verificações. O motor grava prova e intenção antes de mover o original para Itens Excluídos. O movimento registra a nova identidade. Para recuperar uma interrupção em PST, cuja movimentação pode alterar os identificadores, a origem recebe um identificador único de recuperação após a intenção validada e antes do movimento; o conteúdo observado é conferido novamente.

No modo `deleted-items`, o fluxo termina com o original rastreado em Itens Excluídos. No modo `mailbox`, o motor faz nova conferência, grava intenção de remoção final e chama Delete somente nesse item. Nunca esvazia Itens Excluídos nem purga conteúdo recuperável do servidor. Uma ausência após interrupção só pode ser reconciliada com intenção durável de remoção final e cópia/backup atuais íntegros; ausência sem essa evidência vai para revisão.

Para origem Exchange/Microsoft 365, esta versão bloqueia remoção offline ou em cache. A exigência on-line sem cache reduz o risco de validar somente uma cópia local desatualizada, mas não transforma o resultado em certificação do servidor. Retenção e itens recuperáveis podem continuar consumindo espaço. Para origem PST local, o caminho também precisa ser local e fora do OneDrive.

Não edite a origem ou os PSTs da execução durante a operação. ChangeKey é conferido quando disponível; não existe transação distribuída entre SQLite e Outlook. Falhas e ambiguidades preservam evidência e bloqueiam a ação afetada.

## Estados

| Estado | Interpretação |
| --- | --- |
| `Planned`, `Exported`, `ImportStarted` | Cópia ainda não concluída; pode retomar no mesmo catálogo. |
| `Copied` | Cópia registrada, validação ainda pendente. |
| `Verified` | Provas exigidas passaram naquele momento; remoção ainda requer etapa explícita e nova conferência. |
| `DeleteStarted`, `InDeletedItems`, `PurgeStarted` | Etapas de remoção rastreadas; a retomada considera a fase durável. |
| `Removed` | Original rastreado removido/reconciliado segundo as provas do motor. Não certifica quota liberada. |
| `ReviewRequired` | Divergência ou ambiguidade; não há opção para ignorar a prova. |
| `SkippedDraft`, `SkippedUnsupported` | Rascunho ou classe pendente; original preservado. |

## Comandos do motor

```text
archive-ui-sources
archive-plan <config.json>
archive-run <raiz> <execucao>
archive-verify <raiz> <execucao>
archive-reconcile <raiz> <execucao>
archive-remove <raiz> <execucao> deleted-items|mailbox
archive-status <raiz> [execucao]
archive-self-test
```

`archive-run` copia/retoma e depois valida. `archive-status` é consulta local somente leitura; inclui origem, estados, partes e erros. Os demais comandos operacionais conectam ao Outlook. Código 0 indica conclusão do comando sem pendências contempladas pelo seu retorno; 3 sinaliza pendências ou revisão; 1 falha de execução e 2 uso inválido. Consulte sempre os estados, pois conclusão de cópia não é autorização de remoção.

Configuração mínima para uso técnico, com IDs obtidos por `archive-ui-sources`:

```json
{
  "RootDirectory": "C:\\ArquivoOutlook\\colaborador",
  "SourceStoreId": "STORE-ID",
  "SourceFolderId": "FOLDER-ID",
  "SourceLabel": "Caixa > pasta",
  "IncludeSubfolders": true
}
```

A raiz precisa ser absoluta, própria, em disco local fixo, fora do OneDrive/redirecionamentos e com até 120 caracteres. A configuração fica congelada no catálogo ao planejar. Para retomar, use a raiz e execução existentes; não edite o plano ou o banco manualmente.

## Evidências e limites dos testes

A pasta de trabalho guarda catálogo, eventos, plano, partes PST, MSGs e hashes/provas por item. Contém dados de e-mail e deve ser protegida e preservada inteira. Fechar a entrada do PST no Outlook não apaga o arquivo. Não publique relatórios com identificadores da caixa ou conteúdo real.

Os testes automáticos exercitam SQLite/arquivos e provedor artificial; os testes de integração usam cinco mensagens artificiais locais com Unicode, HTML/CID, RTF e anexos de até 2 MiB. Não enviam mensagens. Nenhum original real da empresa foi removido no desenvolvimento. Os testes não homologam caixas completas de 10–50 GB, rascunhos, itens protegidos, classes especiais, preservação de `Hidden=true` ou aparência visual do conteúdo. Resultados específicos da integração são mantidos em relatórios separados.
