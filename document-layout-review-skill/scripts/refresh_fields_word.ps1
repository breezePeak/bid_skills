param(
  [Parameter(Mandatory=$true)][string]$InputPath,
  [Parameter(Mandatory=$true)][string]$OutputPath,
  [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$word = $null; $doc = $null; $oldUpdateLinks = $null
$result = @{status='failed'; engine='Microsoft Word'; fields_updated=$false; indexes_updated=$false; errors=@()}
try {
  $inputFull = [System.IO.Path]::GetFullPath($InputPath)
  $outputFull = [System.IO.Path]::GetFullPath($OutputPath)
  if ($inputFull -eq $outputFull) { throw 'Input and output must differ.' }
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3 # disable document macros
  $oldUpdateLinks = $word.Options.UpdateLinksAtOpen
  $word.Options.UpdateLinksAtOpen = $false
  $doc = $word.Documents.Open($inputFull, $false, $false, $false)
  for ($pass=0; $pass -lt 2; $pass++) {
    foreach ($story in $doc.StoryRanges) {
      $range = $story
      while ($null -ne $range) {
        foreach ($field in $range.Fields) {
          $command = ($field.Code.Text.Trim() -split '\s+')[0].ToUpperInvariant()
          if ($command -in @('DDE','DDEAUTO','INCLUDETEXT','INCLUDEPICTURE','DATABASE','LINK','RD')) {
            throw "External-content field needs explicit review: $command"
          }
          if (-not $field.Update()) { throw "Word reported a field update failure: $command" }
          if ($field.Result.Text -match '错误[！!]\s*(未定义样式|未找到引用源|未定义书签)|Error!\s*(No text of specified style|Reference source not found|Bookmark not defined)') {
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
  $doc.SaveAs2($outputFull, 16)
  $result.status = 'passed'; $result.fields_updated = $true; $result.indexes_updated = $true
  $result.word_version = $word.Version
} catch {
  $result.errors = @($_.Exception.Message)
} finally {
  if ($null -ne $doc) { $doc.Close(0); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
  if ($null -ne $word) {
    if ($null -ne $oldUpdateLinks) { $word.Options.UpdateLinksAtOpen = $oldUpdateLinks }
    $word.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
  $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
}
if ($result.status -ne 'passed') { exit 2 }
