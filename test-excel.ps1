param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$frameworkPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
if (-not (Test-Path (Join-Path $frameworkPath 'csc.exe'))) { $frameworkPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319' }
$compilerPath = Join-Path $frameworkPath 'csc.exe'
$libraryPath = Join-Path $projectRoot 'build/MelloCharts.AddIn.dll'
if (-not (Test-Path $libraryPath)) { throw 'Compile primeiro com ./build.ps1 -RunTests.' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $projectRoot 'build/excel-smoke' }
$testOutput = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $testOutput | Out-Null
$testPath = Join-Path $projectRoot 'build/ExcelSmoke.exe'
& $compilerPath /nologo /target:exe /platform:anycpu "/out:$testPath" /r:System.dll /r:System.Core.dll /r:Microsoft.CSharp.dll "/r:$libraryPath" (Join-Path $projectRoot 'tests/ExcelSmoke.cs')
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o ensaio de Excel.' }
Write-Host 'O ensaio cria um workbook separado por COM. Não instala nem registra o suplemento; fecha somente o workbook de teste.'
& $testPath $testOutput
if ($LASTEXITCODE -ne 0) { throw 'O ensaio do Excel não foi concluído. Confira a mensagem de erro acima.' }
