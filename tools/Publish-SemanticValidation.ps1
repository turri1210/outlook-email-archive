param(
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [string]$DependenciesDirectory=(Join-Path $env:USERPROFILE 'Downloads/outlook-relevant-analysis-deps'),
 [string]$PythonExecutable=(Get-Command python -ErrorAction Stop).Source
)
$ErrorActionPreference='Stop'
if(!(Test-Path -LiteralPath $PythonExecutable -PathType Leaf) -or !(Test-Path -LiteralPath $DependenciesDirectory -PathType Container)){throw 'Runtime ou dependências da validação não encontrados.'}
$destination=Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) 'semantic'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
foreach($name in @('semantic_message_worker.py','compare_relevant_message.py','verify_rtf_sample.py','rtf_representation.py','static_css_profile.py','signed_mime_profile.py','requirements-relevant-analysis.txt')){
 Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination (Join-Path $destination $name) -Force
}
New-Item -ItemType Directory -Path (Join-Path $destination 'site-packages') -Force | Out-Null
Get-ChildItem -LiteralPath $DependenciesDirectory -Force | ForEach-Object {Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $destination 'site-packages') -Recurse -Force}
[IO.File]::WriteAllText((Join-Path $destination 'python.path'),[IO.Path]::GetFullPath($PythonExecutable),[Text.UTF8Encoding]::new($false))
Write-Output 'Validação de informações relevantes publicada.'
