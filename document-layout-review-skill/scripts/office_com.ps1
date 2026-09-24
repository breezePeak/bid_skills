param(
  [Parameter(Mandatory=$true)][ValidateSet('word','wps')][string]$Engine,
  [Parameter(Mandatory=$true)][ValidateSet('probe','render','refresh')][string]$Mode,
  [Parameter(Mandatory=$true)][string]$ReportPath,
  [string]$InputPath,
  [string]$OutputPath
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$app = $null; $doc = $null; $owned = $false
$restoreApp = @{}; $restoreOptions = @{}
$name = if ($Engine -eq 'word') {'Microsoft Word'} else {'WPS Writer'}
$progId = if ($Engine -eq 'word') {'Word.Application'} else {'KWPS.Application'}
$result = @{status='failed'; engine=$name; mode=$Mode; fields_updated=$false;
            indexes_updated=$false; errors=@(); error_code='office-worker-failed'}
function Set-AppOption($key, $value) {
  $restoreApp[$key] = $app.$key
  $app.$key = $value
}
function Set-DocumentOption($key, $value) {
  $restoreOptions[$key] = $app.Options.$key
  $app.Options.$key = $value
}
try {
  $app = New-Object -ComObject $progId
  # Some WPS editions share an application instance. Never close user documents.
  if ($app.Documents.Count -ne 0) { throw 'Office returned an instance with open user documents; no changes made.' }
  $owned = $true
  $result.application = [string]$app.Name
  $result.version = [string]$app.Version
  if (($Engine -eq 'word') -and ($result.application -notmatch '(?i)word')) {
    throw 'Word.Application did not return Microsoft Word.'
  }
  if (($Engine -eq 'wps') -and ($result.application -notmatch '(?i)wps|kingsoft|金山')) {
    throw 'KWPS.Application did not return WPS Writer.'
  }
  if ($Mode -ne 'probe') {
    $inputFull = [IO.Path]::GetFullPath($InputPath)
    $outputFull = [IO.Path]::GetFullPath($OutputPath)
    if ($inputFull -eq $outputFull) { throw 'Input and output must differ.' }
    Set-AppOption 'Visible' $false
    Set-AppOption 'DisplayAlerts' 0
    Set-AppOption 'AutomationSecurity' 3
    Set-DocumentOption 'UpdateLinksAtOpen' $false
    # Rendering must not silently refresh fields different from the saved candidate.
    if ($Mode -eq 'render') {
      Set-DocumentOption 'UpdateFieldsAtPrint' $false
      Set-DocumentOption 'UpdateLinksAtPrint' $false
    }
    $readOnly = ($Mode -eq 'render')
    $doc = $app.Documents.Open($inputFull, $false, $readOnly, $false)
    if ($Mode -eq 'refresh') {
      for ($pass=0; $pass -lt 2; $pass++) {
        foreach ($story in $doc.StoryRanges) {
          $range = $story
          while ($null -ne $range) {
            foreach ($field in $range.Fields) {
              $command = ($field.Code.Text.Trim() -split '\s+')[0].ToUpperInvariant()
              if ($command -in @('DDE','DDEAUTO','INCLUDETEXT','INCLUDEPICTURE','DATABASE','LINK','RD')) {
                $result.error_code='external-field-update-blocked'
                throw "External-content field requires review: $command"
              }
              $updated = $field.Update()
              if ($updated -ne $true) {
                $result.error_code='field-result-error'
                throw "Field update did not confirm success: $command"
              }
              if ($field.Result.Text -match '错误[！!]\s*(未定义样式|未找到引用源|未定义书签)|Error!\s*(No text of specified style|Reference source not found|Bookmark not defined)') {
                $result.error_code='field-result-error'
                throw ('Field result error: ' + $field.Result.Text)
              }
            }
            $range = $range.NextStoryRange
          }
        }
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        foreach ($tof in $doc.TablesOfFigures) { $tof.Update() }
        $doc.Repaginate()
      }
      if ($Engine -eq 'word') { $doc.SaveAs2($outputFull, 16) }
      else { $doc.SaveAs($outputFull, 12) }
      $result.fields_updated=$true; $result.indexes_updated=$true
    } else {
      $doc.Repaginate()
      $doc.ExportAsFixedFormat($outputFull, 17)
      $result.pdf_exported=$true
    }
    if (-not (Test-Path -LiteralPath $outputFull -PathType Leaf)) { throw 'Office produced no output file.' }
    if ((Get-Item -LiteralPath $outputFull).Length -eq 0) { throw 'Office produced an empty output file.' }
  }
  $result.status='passed'; $result.error_code=$null
} catch {
  $result.errors=@($_.Exception.Message)
} finally {
  # Cleanup failures must still produce a diagnostic report. Never taskkill by image name.
  if ($null -ne $doc) {
    try { $doc.Close(0) } catch { $result.status='failed'; $result.errors+=('Close: '+$_.Exception.Message) }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch {}
  }
  if ($null -ne $app) {
    if ($owned) {
      foreach ($key in $restoreOptions.Keys) { try { $app.Options.$key=$restoreOptions[$key] } catch {} }
      foreach ($key in $restoreApp.Keys) { try { $app.$key=$restoreApp[$key] } catch {} }
      try {
        if ($app.Documents.Count -eq 0) { $app.Quit() }
        else { $result.status='failed'; $result.errors+=('Office now contains another document; application left open.') }
      } catch { $result.status='failed'; $result.errors+=('Quit: '+$_.Exception.Message) }
    }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
  }
  $parent = Split-Path -Parent ([IO.Path]::GetFullPath($ReportPath))
  [void][IO.Directory]::CreateDirectory($parent)
  $result | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
}
if ($result.status -ne 'passed') { exit 2 }
