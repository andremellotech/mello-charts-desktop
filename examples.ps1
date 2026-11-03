param(
  [switch]$Verify,
  [string]$Destination = $PSScriptRoot,
  [string]$ArchivePath = (Join-Path $PSScriptRoot 'examples/MelloCharts-Examples.zip')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Get-BytesHash([byte[]]$bytes) {
  $sha = [Security.Cryptography.SHA256]::Create()
  try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant() } finally { $sha.Dispose() }
}
function Read-Entry($entry) {
  if ($entry.Length -gt 10485760) { throw 'Arquivo de exemplo acima do limite.' }
  $stream = $entry.Open(); $memory = New-Object IO.MemoryStream
  try { $stream.CopyTo($memory); return ,$memory.ToArray() } finally { $stream.Dispose(); $memory.Dispose() }
}
$zip = [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($ArchivePath))
try {
  $names = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($entry in $zip.Entries) { if (-not $names.Add($entry.FullName)) { throw 'Nome duplicado no pacote.' } }
  $manifestEntry = $zip.GetEntry('MANIFEST.json')
  if (-not $manifestEntry) { throw 'Manifesto do pacote ausente.' }
  $manifest = [Text.Encoding]::UTF8.GetString((Read-Entry $manifestEntry)) | ConvertFrom-Json
  if ($manifest.version -ne 1 -or @($manifest.files).Count -ne 56 -or $zip.Entries.Count -ne 57) { throw 'Pacote de exemplos incompleto ou com arquivos inesperados.' }
  $expected = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $verified = @()
  foreach ($file in $manifest.files) {
    $name = [string]$file.path
    if (-not $name.StartsWith('examples/', [StringComparison]::Ordinal) -or $name.Contains('\') -or $name.Contains(':') -or $name.Contains([char]0) -or ($name.Split('/') | Where-Object { $_ -in @('', '.', '..') })) { throw 'Caminho inesperado no pacote.' }
    if (-not $expected.Add($name) -or $file.sha256 -notmatch '^[a-f0-9]{64}$') { throw 'Manifesto invalido.' }
    $entry = $zip.GetEntry($name)
    if (-not $entry -or $entry.Length -ne $file.bytes) { throw 'Arquivo ausente ou tamanho alterado: ' + $name }
    $bytes = Read-Entry $entry
    if ((Get-BytesHash $bytes) -ne $file.sha256) { throw 'Hash alterado: ' + $name }
    $verified += [pscustomobject]@{ Name=$name; Bytes=$bytes; Hash=$file.sha256 }
  }
  if ($Verify) { Write-Host '56 exemplos originais verificados por tamanho e SHA-256.'; return }
  $targetRoot = [IO.Path]::GetFullPath($Destination).TrimEnd([IO.Path]::DirectorySeparatorChar)
  if ($targetRoot -eq [IO.Path]::GetPathRoot($targetRoot).TrimEnd([IO.Path]::DirectorySeparatorChar)) { throw 'Escolha uma pasta de destino, nao a raiz do disco.' }
  $prefix = $targetRoot + [IO.Path]::DirectorySeparatorChar
  # Validate every destination before writing any file; preserve local edits.
  foreach ($file in $verified) {
    $target = [IO.Path]::GetFullPath((Join-Path $targetRoot $file.Name))
    if (-not $target.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Caminho fora do destino.' }
    $cursor = $target
    while ($cursor -and $cursor.StartsWith($targetRoot, [StringComparison]::OrdinalIgnoreCase)) {
      if (Test-Path -LiteralPath $cursor) {
        $item = Get-Item -LiteralPath $cursor -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Destino com link de diretorio nao permitido.' }
        if ($cursor -ne $target -and -not $item.PSIsContainer) { throw 'Pasta de destino ocupada por arquivo: ' + $cursor }
      }
      if ($cursor -eq $targetRoot) { break }; $cursor = Split-Path $cursor -Parent
    }
    if (Test-Path -LiteralPath $target) {
      if (-not (Test-Path -LiteralPath $target -PathType Leaf) -or (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash.ToLowerInvariant() -ne $file.Hash) { throw 'Arquivo local diferente; use outra pasta: ' + $target }
    }
  }
  foreach ($file in $verified) {
    $target = Join-Path $targetRoot $file.Name
    if (Test-Path -LiteralPath $target) { continue }
    New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent) | Out-Null
    $stream = [IO.File]::Open($target, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
    try { $stream.Write($file.Bytes, 0, $file.Bytes.Length) } finally { $stream.Dispose() }
  }
  Write-Host "56 exemplos disponiveis em $targetRoot/examples. Arquivos locais diferentes nunca sao sobrescritos."
} finally { $zip.Dispose() }
