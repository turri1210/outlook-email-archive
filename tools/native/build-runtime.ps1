param(
    [Parameter(Mandatory=$true)][string]$ToolchainRoot
)
$ErrorActionPreference='Stop'
$projectRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$compiler=Join-Path ([IO.Path]::GetFullPath($ToolchainRoot)) 'bin/x86_64-w64-mingw32-clang++.exe'
if(!(Test-Path -LiteralPath $compiler -PathType Leaf)){throw 'Compilador portátil x64 não encontrado.'}
# Package asset only. This script never starts Outlook or replaces running DLLs.
$assetDirectory=Join-Path $projectRoot 'src/native/win-x64'
New-Item -ItemType Directory -Path $assetDirectory -Force | Out-Null
& $compiler (Join-Path $PSScriptRoot 'native-copy.cpp') ('-I'+(Join-Path $PSScriptRoot 'mapi-headers')) -static -municode -O2 -Wno-nonportable-include-path -o (Join-Path $assetDirectory 'native-copy.exe')
if($LASTEXITCODE -ne 0){throw 'Falha na compilação do componente nativo de cópia.'}
Get-FileHash -LiteralPath (Join-Path $assetDirectory 'native-copy.exe') -Algorithm SHA256 | Select-Object Path,Hash

& $compiler (Join-Path $PSScriptRoot 'native-identity.cpp') ('-I'+(Join-Path $PSScriptRoot 'mapi-headers')) -static -municode -O2 -Wno-nonportable-include-path -o (Join-Path $assetDirectory 'native-identity.exe')
if($LASTEXITCODE -ne 0){throw 'Falha na compilação do componente nativo de identificação.'}
& python (Join-Path $PSScriptRoot '../test-native-json.py') (Join-Path $assetDirectory 'native-identity.exe')
if($LASTEXITCODE -ne 0){throw 'Falha no teste de escape JSON do componente nativo.'}
Get-FileHash -LiteralPath (Join-Path $assetDirectory 'native-identity.exe') -Algorithm SHA256 | Select-Object Path,Hash

& $compiler (Join-Path $PSScriptRoot 'native-finalize.cpp') ('-I'+(Join-Path $PSScriptRoot 'mapi-headers')) -static -municode -O2 -Wno-nonportable-include-path -o (Join-Path $assetDirectory 'native-finalize.exe')
if($LASTEXITCODE -ne 0){throw 'Falha na compilação do componente nativo de finalização.'}
