# Conferir e corrigir pendências

Selecione a execução original no histórico. Na aba **Pendências**, cada erro traz uma orientação. **Salvar relatório…** exporta as contagens e os erros exibidos, com a chave completa de cada item. O relatório não contém o corpo das mensagens. As contagens incluem itens copiados que aguardam validação, mesmo quando não existe erro individual.

1. **Reconciliar pendências** procura cópias já existentes e tenta registrá-las com segurança. Não repete o envio de um lote. Se a correspondência for ambígua ou o conteúdo não passar pela conferência, o item permanece pendente.
2. **Verificar novamente** executa a conferência detalhada das cópias registradas. Esta etapa pode demorar. Uma transferência rápida em estado **Copiado** ainda precisa passar por esta validação antes de permitir a remoção do original.
3. **Copiar / retomar** continua os itens restantes da lista inicial. Não prepare outro plano para corrigir a mesma execução: isso dificulta a identificação das cópias anteriores.

Se a pendência indicar que o PST está ocupado, feche mensagens ou janelas que estejam exibindo esse arquivo e tente verificar novamente. Preserve a pasta inteira do histórico, o catálogo SQLite e os backups MSG já gerados. Não exclua mensagens copiadas nem altere o catálogo para forçar uma repetição.

Divergências de conteúdo, cópias ambíguas e lotes incompletos não são corrigidos apagando os originais. Compare origem e destino e mantenha os originais enquanto a pendência não tiver sido esclarecida. Rascunhos e tipos de item ainda não suportados permanecem fora da conclusão automática.

As novas mensagens que chegaram depois do planejamento ficam fora da lista inicial. Elas podem originar outro plano após a conclusão, sem substituir os itens em andamento.


## Correção de 28/10/2026

A reconciliação também reavalia mensagens em revisão depois de um lote concluído. Os backups e as evidências originais permanecem imutáveis. A extração HTML descarta marcadores `pntext` de RTF somente quando a estrutura HTML comprova que são repetições dos marcadores da própria lista; texto literal, links e imagens continuam conferidos. Corpos vazios derivados de texto recebem um perfil específico com validação RTF restrita.

Uma origem não localizada e uma cópia sem correspondência no destino previsto são situações diferentes. O registro agora identifica a etapa e o motivo. Não são corrigidas trocando a mensagem por outra ou repetindo a transferência automaticamente. Preserve os MSGs e use o relatório para localizar os itens. A confirmação da gravação do PST em disco continua sendo uma etapa distinta da conferência do conteúdo; a política de remoção dos originais não foi relaxada.

Quando a consulta não encontra novas mensagens, a interface registra apenas: “Não há novos e-mails fora do escopo desta execução.” O JSON interno da consulta deixa de aparecer no registro da interface.

## Escopo confirmado em 30/10/2026

O usuário autorizou desconsiderar convites de reunião, respostas de aceitação e cancelamentos. Esses itens podem ficar fora do arquivamento; isso não autoriza excluir seus originais. Avisos de falha na entrega (REPORT.IPM.Note.NDR) continuam no escopo desejado e precisam de suporte próprio para cópia e validação, pois o caminho atual aceita somente IPM.Note/Class 43.
