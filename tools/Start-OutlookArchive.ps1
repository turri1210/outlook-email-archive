param([switch]$Rebuild)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$uiExecutable = Join-Path $projectRoot 'ui\bin\Release\net8.0-windows\OutlookArchive.UI.exe'
$engineAssembly = Join-Path $projectRoot 'src\bin\Release\net8.0-windows\OutlookArchive.dll'
$requiredComponents = @('native-copy.exe','native-identity.exe','native-finalize.exe')
$missingRuntime = @($requiredComponents | Where-Object { !(Test-Path -LiteralPath (Join-Path (Split-Path -Parent $engineAssembly) $_) -PathType Leaf) })
if ($Rebuild -or $missingRuntime.Count -gt 0 -or !(Test-Path -LiteralPath $uiExecutable) -or !(Test-Path -LiteralPath $engineAssembly)) {
    & dotnet build (Join-Path $projectRoot 'src\OutlookArchive.csproj') -c Release --nologo
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o motor de arquivamento.' }
    & dotnet build (Join-Path $projectRoot 'ui\OutlookArchive.UI.csproj') -c Release --nologo
    if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar a interface.' }
}
# This is the user-facing interactive application requested by the user.
Start-Process -FilePath $uiExecutable -WorkingDirectory $projectRoot
