param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$frameworkPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
if (-not (Test-Path (Join-Path $frameworkPath 'csc.exe'))) { $frameworkPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319' }
$compilerPath = Join-Path $frameworkPath 'csc.exe'
$libraryPath = Join-Path $projectRoot 'build/MelloCharts.AddIn.dll'
if (-not (Test-Path $libraryPath)) { throw 'Compile primeiro com ./build.ps1 -RunTests.' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $projectRoot 'build/office-smoke' }
$testOutput = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $testOutput | Out-Null
$testPath = Join-Path $projectRoot 'build/OfficeSmoke.exe'
& $compilerPath /nologo /target:exe /platform:anycpu "/out:$testPath" /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll "/r:$libraryPath" (Join-Path $projectRoot 'tests/OfficeSmoke.cs')
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o ensaio de Office.' }
Write-Host 'O ensaio cria uma apresentação separada por COM. Não instala nem registra o suplemento.'
& $testPath $testOutput
if ($LASTEXITCODE -ne 0) { throw 'O ensaio do PowerPoint não foi concluído. Confira a mensagem de erro acima.' }
