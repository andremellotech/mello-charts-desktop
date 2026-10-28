param([switch]$RunTests,[switch]$RenderGallery)
$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$buildRoot = Join-Path $projectRoot 'build'
$releaseRoot = Join-Path $projectRoot 'release'
$frameworkPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
if (-not (Test-Path (Join-Path $frameworkPath 'csc.exe'))) { $frameworkPath = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319' }
$compilerPath = Join-Path $frameworkPath 'csc.exe'
if (-not (Test-Path $compilerPath)) { throw 'Instale ou habilite o .NET Framework 4.x para compilar.' }
New-Item -ItemType Directory -Force -Path $buildRoot, $releaseRoot | Out-Null
$references = @('/r:System.dll','/r:System.Core.dll','/r:System.Drawing.dll','/r:System.Windows.Forms.dll','/r:System.Web.Extensions.dll','/r:System.Xml.dll','/r:System.Xml.Linq.dll','/r:System.IO.Compression.dll','/r:System.IO.Compression.FileSystem.dll','/r:Microsoft.CSharp.dll')
$libraryPath = Join-Path $buildRoot 'MelloCharts.AddIn.dll'
$sources = @('src/ChartCore.cs','src/ChartColors.cs','src/ChartColorForm.cs','src/ChartCatalog.cs','src/AdvancedCharts.cs','src/ChartGalleryForm.cs','src/EditorForm.cs','src/PowerPointAddIn.cs','src/NativeIntegration.cs','src/InstallationHealth.cs','src/ProfessionalCharts.cs','src/WorkbookData.cs','src/WorkbookImportForm.cs','src/ExcelAddIn.cs') | ForEach-Object { Join-Path $projectRoot $_ }
& $compilerPath /nologo /target:library /platform:anycpu /optimize+ "/out:$libraryPath" $references "/resource:$(Join-Path $projectRoot 'assets/icon.png'),MelloCharts.Icon.png" $sources
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o suplemento.' }
$setupPath = Join-Path $releaseRoot 'MelloCharts-Setup.exe'
& $compilerPath /nologo /target:winexe /platform:anycpu /optimize+ "/out:$setupPath" "/win32icon:$(Join-Path $projectRoot 'assets/icon.ico')" /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll "/resource:$libraryPath,MelloCharts.AddIn.dll" (Join-Path $projectRoot 'installer/Setup.cs') (Join-Path $projectRoot 'src/InstallationHealth.cs')
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o instalador.' }
$portableRoot = Join-Path $buildRoot 'portable'
New-Item -ItemType Directory -Force -Path $portableRoot | Out-Null
$launcherPath = Join-Path $portableRoot 'MelloCharts.exe'
& $compilerPath /nologo /target:winexe /platform:anycpu /optimize+ "/out:$launcherPath" "/win32icon:$(Join-Path $projectRoot 'assets/icon.ico')" $references "/r:$libraryPath" (Join-Path $projectRoot 'src/DesktopLauncher.cs')
if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar o editor independente.' }
Copy-Item -LiteralPath $libraryPath -Destination (Join-Path $portableRoot 'MelloCharts.AddIn.dll') -Force
Copy-Item -LiteralPath (Join-Path $projectRoot 'LEIA-PRIMEIRO.txt'),(Join-Path $projectRoot 'LICENSE') -Destination $portableRoot -Force
$portableZip = Join-Path $releaseRoot 'MelloCharts-Portable.zip'
Compress-Archive -LiteralPath (Join-Path $portableRoot 'MelloCharts.exe'),(Join-Path $portableRoot 'MelloCharts.AddIn.dll'),(Join-Path $portableRoot 'LEIA-PRIMEIRO.txt'),(Join-Path $portableRoot 'LICENSE') -DestinationPath $portableZip -Force
if ($RunTests) {
  $coreTestPath = Join-Path $buildRoot 'CoreTests.exe'
  & $compilerPath /nologo /target:exe /platform:anycpu "/out:$coreTestPath" $references "/r:$libraryPath" (Join-Path $projectRoot 'tests/CoreTests.cs')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar testes do núcleo.' }
  & $coreTestPath
  if ($LASTEXITCODE -ne 0) { throw 'Testes do núcleo falharam.' }
  $bridgeTestPath = Join-Path $buildRoot 'BridgeTests.exe'
  & $compilerPath /nologo /target:exe /platform:anycpu "/out:$bridgeTestPath" $references "/r:$libraryPath" (Join-Path $projectRoot 'tests/BridgeTests.cs')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar testes da ponte.' }
  & $bridgeTestPath
  if ($LASTEXITCODE -ne 0) { throw 'Testes da ponte falharam.' }
  foreach ($catalogTestName in @('ExtendedChartTests','EditorCatalogTests','ColorTests','ProfessionalChartTests','WorkbookDataTests','WorkbenchTests','ExcelIntegrationTests')) {
    $catalogTestPath = Join-Path $buildRoot ($catalogTestName + '.exe')
    $catalogTestSource = Join-Path $projectRoot ('tests/' + $catalogTestName + '.cs')
    & $compilerPath /nologo /target:exe /platform:anycpu "/out:$catalogTestPath" $references "/r:$libraryPath" $catalogTestSource
    if ($LASTEXITCODE -ne 0) { throw "Falha ao compilar $catalogTestName." }
    & $catalogTestPath
    if ($LASTEXITCODE -ne 0) { throw "$catalogTestName falhou." }
  }
  $integrationTestPath = Join-Path $buildRoot 'IntegrationTests.exe'
  & $compilerPath /nologo /target:exe /platform:anycpu /main:IntegrationTests "/out:$integrationTestPath" $references "/r:$libraryPath" (Join-Path $projectRoot 'tests/BridgeTests.cs') (Join-Path $projectRoot 'tests/IntegrationTests.cs')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar testes de integração.' }
  & $integrationTestPath
  if ($LASTEXITCODE -ne 0) { throw 'Testes de integração falharam.' }
}
if ($RenderGallery) {
  $galleryToolPath = Join-Path $buildRoot 'RenderGallery.exe'
  & $compilerPath /nologo /target:exe /platform:anycpu "/out:$galleryToolPath" $references "/r:$libraryPath" (Join-Path $projectRoot 'tests/RenderGallery.cs')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar a galeria.' }
  & $galleryToolPath (Join-Path $projectRoot 'examples/gallery')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao renderizar a galeria.' }
  $colorRenderPath = Join-Path $buildRoot 'RenderColors.exe'
  & $compilerPath /nologo /target:exe /platform:anycpu "/out:$colorRenderPath" $references "/r:$libraryPath" (Join-Path $projectRoot 'tests/RenderColors.cs')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar os exemplos de cores.' }
  & $colorRenderPath (Join-Path $projectRoot 'examples/gallery')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao renderizar os exemplos de cores.' }
  $workbenchPath = Join-Path $buildRoot 'RenderWorkbench.exe'
  & $compilerPath /nologo /target:exe /platform:anycpu "/out:$workbenchPath" $references "/r:$libraryPath" (Join-Path $projectRoot 'tests/RenderWorkbench.cs')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao compilar prévias do editor.' }
  & $workbenchPath (Join-Path $projectRoot 'examples/gallery')
  if ($LASTEXITCODE -ne 0) { throw 'Falha ao renderizar prévias do editor.' }
}
$setupHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $setupPath).Hash.ToLowerInvariant()
$portableHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $portableZip).Hash.ToLowerInvariant()
[IO.File]::WriteAllText((Join-Path $releaseRoot 'SHA256SUMS.txt'), "$setupHash  MelloCharts-Setup.exe`n$portableHash  MelloCharts-Portable.zip`n", [Text.Encoding]::ASCII)
Get-FileHash -Algorithm SHA256 -LiteralPath $setupPath | Format-List
Write-Host "Instalador gerado: $setupPath"
Write-Host 'Este script compila e testa. Não instala, registra nem abre o PowerPoint.'
