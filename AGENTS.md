# Entregas deste projeto

Para compilar uma versão destinada ao usuário, execute `tools/Publish-Downloads.ps1`.
O script compila, testa e atualiza juntos o programa e `Arquivo Outlook.zip` em
`Downloads/emailArchive`. Toda entrega deve atualizar esse ZIP.

Não substitua os binários enquanto a interface ou o motor da instalação estiverem
em execução. Compilações isoladas de desenvolvimento/testes ficam em `runs`.
O ZIP compartilhável deve conter apenas o programa e dependências, sem a pasta
de dados `emailArchive`, PSTs, MSGs, catálogo, logs ou configurações pessoais.
