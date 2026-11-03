$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path $PSScriptRoot -Parent
$helper = Join-Path $projectRoot 'examples.ps1'
$archivePath = Join-Path $projectRoot 'examples/MelloCharts-Examples.zip'
if (-not (Test-Path -LiteralPath $helper)) { throw 'Extrator/verificador de exemplos ausente.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('mello-examples-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
$checks = 0
function Assert-Condition($condition, $message) { if (-not $condition) { throw $message }; $script:checks++ }
function Assert-Fails([scriptblock]$action, $message) { $failed = $false; try { & $action | Out-Null } catch { $failed = $true }; Assert-Condition $failed $message }
try {
  & $helper -Verify
  $target = Join-Path $testRoot 'restored'
  & $helper -Destination $target
  Assert-Condition ((Get-ChildItem -LiteralPath (Join-Path $target 'examples') -Recurse -File).Count -eq 56) 'Pacote deve restaurar os 56 arquivos originais.'
  & $helper -Destination $target
  $original = Join-Path $target 'examples/buildup.csv'
  [IO.File]::WriteAllText($original, 'edicao local')
  Assert-Fails { & $helper -Destination $target } 'Edicao local nao pode ser sobrescrita.'
  Assert-Condition ([IO.File]::ReadAllText($original) -eq 'edicao local') 'Edicao local foi alterada.'
  $conflict = Join-Path $testRoot 'conflict'
  New-Item -ItemType Directory -Path (Join-Path $conflict 'examples') -Force | Out-Null
  [IO.File]::WriteAllText((Join-Path $conflict 'examples/waterfall.json'), 'preservar')
  Assert-Fails { & $helper -Destination $conflict } 'Conflito deve falhar antes de extrair outros exemplos.'
  Assert-Condition ((Get-ChildItem -LiteralPath $conflict -Recurse -File).Count -eq 1) 'Extracao parcial antes de validar todos os conflitos.'
  $blocked = Join-Path $testRoot 'blocked'
  New-Item -ItemType Directory -Path (Join-Path $blocked 'examples') -Force | Out-Null
  [IO.File]::WriteAllText((Join-Path $blocked 'examples/gallery'), 'arquivo local')
  Assert-Fails { & $helper -Destination $blocked } 'Pasta ocupada por arquivo deve falhar na validacao inicial.'
  Assert-Condition ((Get-ChildItem -LiteralPath $blocked -Recurse -File).Count -eq 1) 'Extracao parcial diante de pasta ocupada por arquivo.'
  $broken = Join-Path $testRoot 'broken.zip'
  Copy-Item -LiteralPath $archivePath -Destination $broken
  $zip = [IO.Compression.ZipFile]::Open($broken, [IO.Compression.ZipArchiveMode]::Update)
  try {
    $zip.GetEntry('examples/buildup.csv').Delete()
    $entry = $zip.CreateEntry('examples/buildup.csv')
    $writer = New-Object IO.StreamWriter($entry.Open()); try { $writer.Write('alterado') } finally { $writer.Dispose() }
  } finally { $zip.Dispose() }
  Assert-Fails { & $helper -ArchivePath $broken -Verify } 'Arquivo alterado sem atualizar hash deve ser rejeitado.'
  $slip = Join-Path $testRoot 'slip.zip'
  Copy-Item -LiteralPath $archivePath -Destination $slip
  $zip = [IO.Compression.ZipFile]::Open($slip, [IO.Compression.ZipArchiveMode]::Update)
  try { $entry = $zip.CreateEntry('../outside.txt'); $entry.Open().Dispose() } finally { $zip.Dispose() }
  Assert-Fails { & $helper -ArchivePath $slip -Destination (Join-Path $testRoot 'slip-output') } 'Caminho inesperado deve ser rejeitado.'
  Assert-Condition (-not (Test-Path -LiteralPath (Join-Path $testRoot 'outside.txt'))) 'Pacote escreveu fora do destino.'
  Write-Host "$checks verificacoes do pacote de exemplos aprovadas."
} finally {
  $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
  $allowedRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
  if (-not $resolvedTestRoot.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Limpeza fora da pasta temporaria.' }
  Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
}
