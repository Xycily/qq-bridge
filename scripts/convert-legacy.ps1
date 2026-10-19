# convert-legacy.ps1 - convert legacy Office formats (.doc/.xls/.ppt) to modern ones via COM.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File convert-legacy.ps1 -InPath a.doc -OutPath a.docx
#
# Notes:
# - Tries Microsoft Office first (Word/Excel/PowerPoint.Application), then WPS (KWps/KET/KWPP).
# - No windows, no dialogs. Exit code 0 means success.
# - Open -> SaveAs -> Close. The source file is never modified.
# - Keep this file ASCII-only: Windows PowerShell 5.1 reads .ps1 as ANSI when there is no BOM,
#   and non-ASCII comment bytes can break parsing.
param(
  [Parameter(Mandatory = $true)][string]$InPath,
  [Parameter(Mandatory = $true)][string]$OutPath
)

$ErrorActionPreference = 'Stop'

$in = (Resolve-Path -LiteralPath $InPath).Path
$outFull = [System.IO.Path]::GetFullPath($OutPath)
$dir = [System.IO.Path]::GetDirectoryName($outFull)
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
if (Test-Path -LiteralPath $outFull) { Remove-Item -LiteralPath $outFull -Force -ErrorAction SilentlyContinue }

$ext = [System.IO.Path]::GetExtension($in).ToLowerInvariant()

function Try-Word([string]$progId, [int]$format) {
  $app = $null
  $doc = $null
  try {
    $app = New-Object -ComObject $progId
    $app.Visible = $false
    try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Open($in, $false, $true)
    try { $doc.SaveAs2($outFull, $format) } catch { $doc.SaveAs($outFull, $format) }
    return $true
  } catch {
    return $false
  } finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    if ($app) { try { $app.Quit() } catch {} }
  }
}

function Try-Excel([string]$progId, [int]$format) {
  $app = $null
  $wb = $null
  try {
    $app = New-Object -ComObject $progId
    try { $app.Visible = $false } catch {}
    try { $app.DisplayAlerts = $false } catch {}
    $wb = $app.Workbooks.Open($in, 0, $true)
    $wb.SaveAs($outFull, $format)
    return $true
  } catch {
    return $false
  } finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($app) { try { $app.Quit() } catch {} }
  }
}

function Try-PowerPoint([string]$progId, [int]$format) {
  $app = $null
  $pres = $null
  try {
    $app = New-Object -ComObject $progId
    $pres = $app.Presentations.Open($in, $true, $false, $false)
    $pres.SaveAs($outFull, $format)
    return $true
  } catch {
    return $false
  } finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($app) { try { $app.Quit() } catch {} }
  }
}

$ok = $false
# Measured on this machine: WPS COM responds quickly and reliably, while Microsoft Word
# tends to hang on its start screen in a non-interactive session. So try WPS first.
switch ($ext) {
  '.doc' { $ok = (Try-Word 'KWps.Application' 16) -or (Try-Word 'Word.Application' 16) }
  '.xls' { $ok = (Try-Excel 'KET.Application' 51) -or (Try-Excel 'Excel.Application' 51) }
  '.ppt' { $ok = (Try-PowerPoint 'KWPP.Application' 24) -or (Try-PowerPoint 'PowerPoint.Application' 24) }
  default { Write-Error "unsupported extension: $ext"; exit 2 }
}

if (-not $ok -or -not (Test-Path -LiteralPath $outFull)) {
  Write-Error "conversion failed (Office/WPS missing or file locked): $in"
  exit 1
}

if ((Get-Item -LiteralPath $outFull).Length -le 0) {
  Write-Error "empty output: $outFull"
  exit 1
}

Write-Output "OK $outFull"
exit 0
