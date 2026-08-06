param(
  [string]$JobFile
)

$ErrorActionPreference = 'Stop'

$jobs = Get-Content -Raw -Path $JobFile | ConvertFrom-Json
$total = $jobs.Count
$i = 0

$word = $null
try {
  $word = New-Object -ComObject Word.Application
  # Visible = $true a proposito, aunque la ventana moleste: al abrir un PDF, Word
  # muestra el dialogo de "PDF Reflow", que DisplayAlerts = 0 no suprime. Con la
  # ventana oculta el dialogo existe igual pero nadie puede contestarlo y la
  # llamada COM queda colgada para siempre (probado: minimizarla tampoco alcanza).
  $word.Visible = $true
  $word.DisplayAlerts = 0
} catch {
  Write-Output "FATAL:No se pudo iniciar Word: $($_.Exception.Message)"
  exit 1
}

$results = @()

foreach ($job in $jobs) {
  $i++
  # ${i} y no $i: PowerShell lee "$i:" como variable con unidad (estilo $env:PATH)
  # y el script entero deja de parsear.
  Write-Output "PROGRESS:${i}:${total}:$($job.input)"

  $doc = $null
  try {
    # Al abrir un .pdf, Word dispara su conversión automática ("PDF Reflow")
    # y lo abre como documento editable.
    $doc = $word.Documents.Open($job.input, $false, $false)
    # wdFormatXMLDocument = 12  (formato .docx)
    $doc.SaveAs2($job.output, 12)
    $doc.Close([ref]$false)
    $results += [PSCustomObject]@{ file = $job.input; success = $true; error = $null }
  } catch {
    if ($doc) {
      try { $doc.Close([ref]$false) } catch {}
    }
    $results += [PSCustomObject]@{ file = $job.input; success = $false; error = $_.Exception.Message }
  }
}

try {
  $word.Quit()
} catch {}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$json = ConvertTo-Json $results -Compress
Write-Output "RESULT:$json"
