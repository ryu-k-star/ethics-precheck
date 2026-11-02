param(
  [Parameter(Mandatory=$true)][string]$InputDocx,
  [Parameter(Mandatory=$true)][string]$OutputPdf
)
$resolvedInput = (Resolve-Path -LiteralPath $InputDocx).Path
$outputParent = Split-Path -Parent $OutputPdf
if (-not (Test-Path -LiteralPath $outputParent)) {
  New-Item -ItemType Directory -Path $outputParent | Out-Null
}
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPdf)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($resolvedInput, $false, $true)
  try {
    $doc.ExportAsFixedFormat($resolvedOutput, 17)
  } finally {
    $doc.Close($false)
  }
} finally {
  $word.Quit()
}
Write-Output $resolvedOutput
