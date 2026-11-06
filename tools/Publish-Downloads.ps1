#Requires -Version 7.0
param(
    [string]$Destination=(Join-Path $env:USERPROFILE 'Downloads/emailArchive')
)
$ErrorActionPreference='Stop'
$project=Split-Path -Parent $PSScriptRoot
$destinationPath=[IO.Path]::GetFullPath($Destination)
function Assert-Closed {
    $active=@(Get-CimInstance Win32_Process | Where-Object {
        $_.Name -in @('OutlookArchive.UI.exe','dotnet.exe','OutlookArchive.exe') -and
        $_.CommandLine -and $_.CommandLine.Replace('/','\').Contains($destinationPath.Replace('/','\'),[StringComparison]::OrdinalIgnoreCase)
    })
    if($active.Count){throw 'Feche a interface e aguarde a operação terminar antes de atualizar o programa e o ZIP.'}
}
Assert-Closed
$stage=Join-Path $project ('runs/distribution-'+[Guid]::NewGuid().ToString('N'))
$payload=Join-Path $stage 'program'
$ui=Join-Path $stage 'ui'
$engine=Join-Path $payload 'engine'
New-Item -ItemType Directory -Path $payload,$ui,$engine -Force | Out-Null
& dotnet build (Join-Path $project 'src/OutlookArchive.csproj') -c Release -o $engine --nologo
if($LASTEXITCODE -ne 0){throw 'Falha ao compilar o motor.'}
& dotnet build (Join-Path $project 'ui/OutlookArchive.UI.csproj') -c Release -o $ui --nologo
if($LASTEXITCODE -ne 0){throw 'Falha ao compilar a interface.'}
& dotnet (Join-Path $engine 'OutlookArchive.dll') archive-planning-self-test > (Join-Path $stage 'planning-tests.log') 2>&1
if($LASTEXITCODE -ne 0){throw ('Falha nos testes de retomada do planejamento. Consulte '+$stage)}
& dotnet (Join-Path $engine 'OutlookArchive.dll') archive-transport-self-test > (Join-Path $stage 'transport-tests.log') 2>&1
if($LASTEXITCODE -ne 0){throw ('Falha nos testes de seleção do transporte. Consulte '+$stage)}
$test=Start-Process -FilePath (Join-Path $ui 'OutlookArchive.UI.exe') -ArgumentList '--self-test' -WindowStyle Hidden -Wait -PassThru -RedirectStandardOutput (Join-Path $stage 'ui-tests.log') -RedirectStandardError (Join-Path $stage 'ui-tests-errors.log')
if($test.ExitCode -ne 0){throw ('Falha nos testes da interface. Consulte '+$stage)}
Get-ChildItem -LiteralPath $ui -File | Where-Object Extension -ne '.pdb' | Copy-Item -Destination $payload
$semantic=Join-Path $destinationPath 'engine/semantic'
if(!(Test-Path -LiteralPath (Join-Path $semantic 'site-packages'))){throw 'Dependências da validação ausentes. Prepare engine/semantic com Publish-SemanticValidation.ps1 antes da publicação.'}
Copy-Item -LiteralPath $semantic -Destination $engine -Recurse
foreach($name in @('semantic_message_worker.py','compare_relevant_message.py','verify_rtf_sample.py','rtf_representation.py','static_css_profile.py','signed_mime_profile.py','requirements-relevant-analysis.txt')){
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $engine 'semantic') -Force
}
$readme=@'
Arquivo de e-mails — Outlook clássico

Extraia todos os arquivos do ZIP para uma pasta local antes de abrir OutlookArchive.UI.exe.
Requisitos: Windows x64, Outlook clássico aberto, .NET 8 Desktop Runtime e Python 3.14.
O Python deve estar em %LOCALAPPDATA%\Programs\Python\Python314\python.exe.
Se estiver em outro local, crie engine\semantic\python.path com o caminho completo do python.exe.

Use Continuar arquivamento para processar os lotes pendentes.
Pausar após este lote conclui o lote atual; Interromper preserva os checkpoints.
O movimento automático autorizado leva apenas originais validados para Itens Excluídos.

O PST fica no destino escolhido. Os dados de trabalho e logs ficam em emailArchive, ao lado do programa.
O ZIP contém somente o programa e suas dependências; não inclui e-mails, PSTs, catálogo ou histórico.
Ao atualizar uma instalação, preserve sua pasta emailArchive e os arquivos PST.
'@
[IO.File]::WriteAllText((Join-Path $payload 'LEIA-ME.txt'),$readme,[Text.UTF8Encoding]::new($false))
Add-Type -AssemblyName System.IO.Compression
$zipPath=Join-Path $stage 'Arquivo Outlook.zip'
$archive=[IO.Compression.ZipFile]::Open($zipPath,[IO.Compression.ZipArchiveMode]::Create)
try{
    foreach($file in Get-ChildItem -LiteralPath $payload -Recurse -File){
        $relative=[IO.Path]::GetRelativePath($payload,$file.FullName).Replace('\','/')
        if($relative -match '(^|/)(__pycache__|emailArchive)(/|$)' -or $file.Extension -in @('.pdb','.pyc','.pst','.ost','.msg','.eml','.sqlite','.db','.log') -or $file.Name -eq 'python.path'){continue}
        [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive,$file.FullName,$relative,[IO.Compression.CompressionLevel]::Optimal)
    }
}finally{$archive.Dispose()}
Assert-Closed
New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
Get-ChildItem -LiteralPath $payload | Copy-Item -Destination $destinationPath -Recurse -Force
Copy-Item -LiteralPath $zipPath -Destination (Join-Path $destinationPath 'Arquivo Outlook.zip') -Force
Write-Output ('Programa e ZIP atualizados em '+$destinationPath)
Write-Output ('Testes e pacote preservados em '+$stage)
