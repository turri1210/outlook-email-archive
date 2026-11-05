# Experimento nativo de cópia em lote

Este diretório contém um **experimento artificial**, sem integração com `archive-run`, sem remoção e sem aprovação para produção. O helper C++ usa Extended MAPI em código nativo, sem interop .NET para essas interfaces. Os sinalizadores de `CopyMessages` são zero: somente cópia. Não há `Delete`, `Send`, `MESSAGE_MOVE` ou `SetReadFlags`.

## Dependências e procedência

- Compilador portátil: [llvm-mingw](https://github.com/mstorsjo/llvm-mingw), distribuição oficial `llvm-mingw-20261020-ucrt-x86_64.zip`. Sem instalação ou alteração do sistema. Prefira manter a ferramenta fora do OneDrive; o script recebe seu caminho explicitamente.
- Cabeçalhos `mapi-headers`: [Microsoft MAPIStubLibrary](https://github.com/microsoft/MAPIStubLibrary/tree/main/include), obtidos em 27/10/2026. A licença original MIT está preservada em `mapi-headers/LICENSE`. Esses arquivos de terceiros não foram modificados.
- Documentação: [IMAPIFolder::CopyMessages](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/imapifolder-copymessages). Uma falha pode deixar cópias parciais; o experimento nunca repete automaticamente um lote com erro.

## Compilação isolada

```powershell
./tools/native/build-probe.ps1 -ToolchainRoot 'C:\Users\Usuario\Downloads\outlook-native-toolchain-20261027\llvm-mingw-20261020-ucrt-x86_64'
```

A saída deve ficar dentro de `runs`. O script não substitui os binários de produção. O helper e o Outlook clássico devem ter a mesma arquitetura x64.

## Ensaio artificial

```powershell
./runs/native-preview/OutlookArchive.exe native-batch-probe 'C:\Users\Usuario\Downloads\outlook-local-fixtures-20261027-08\manifest.json' 'C:\Users\Usuario\Downloads\outlook-native-batch-NOVO' 100
```

O manifesto deve indicar o conjunto original de cinco mensagens artificiais, com destinatário `fixture@example.invalid`, em `source-fixtures.pst`. O diretório de saída deve ser novo e local, fora do OneDrive. O argumento opcional é a quantidade de lotes de cinco (1–20) ou `100`, para preparar uma nova origem artificial e testar um único lote de cem itens. As cem mensagens têm cinco conteúdos repetidos vinte vezes; a conferência compara o multiconjunto integral dos hashes, preservando a multiplicidade de mensagens idênticas.

O ensaio grava snapshots selecionados de metadados, streams nativos do corpo e anexos antes/depois; registra separadamente tempo dentro de `CopyMessages`, tempo do processo com inicialização/logon e tempo de conferência. Também mede `Export`, `Import` e `TargetSnapshot` do backend MSG atual sobre cinco mensagens em outro PST novo. O baseline não inclui o catálogo, a retomada ou a verificação final completa do motor.

Todos os PSTs ficam anexados ao terminar para preservar evidências. Remova do perfil somente os caminhos artificiais exatos após inspeção. Nunca faça limpeza por nome genérico “Outlook Data File”.
