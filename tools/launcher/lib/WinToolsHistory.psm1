function Write-WinToolsHistory {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)]$Record,
    [string]$Directory = (Join-Path $env:LOCALAPPDATA 'WinTools\History')
  )
  [void][IO.Directory]::CreateDirectory($Directory)
  $path = Join-Path $Directory ((Get-Date -Format 'yyyy-MM') + '.history.jsonl')
  $line = ($Record | ConvertTo-Json -Depth 5 -Compress) + "`r`n"
  $bytes = [Text.UTF8Encoding]::new($false).GetBytes($line)
  # Exclusive writer access prevents simultaneous launchers from interleaving records.
  for ($attempt = 0; $attempt -lt 5; $attempt++) {
    $stream = $null
    try {
      $stream = [IO.File]::Open($path, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::Read)
      $stream.Write($bytes, 0, $bytes.Length)
      return
    } catch [IO.IOException] {
      if ($attempt -eq 4) { throw }
    } finally {
      if ($stream) { $stream.Dispose() }
    }
    Start-Sleep -Milliseconds 50
  }
}

function Get-WinToolsHistory {
  [CmdletBinding()]
  param(
    [ValidateRange(1, 10000)][int]$Last = 50,
    [string]$Directory = (Join-Path $env:LOCALAPPDATA 'WinTools\History')
  )
  if (-not (Test-Path -LiteralPath $Directory)) { return }
  $remaining = $Last
  $files = @(Get-ChildItem -LiteralPath $Directory -Filter '*.history.jsonl' -File -ErrorAction Stop | Sort-Object Name -Descending)
  foreach ($file in $files) {
    $records = New-Object 'Collections.Generic.Queue[object]'
    $invalid = 0
    Get-Content -LiteralPath $file.FullName -Encoding UTF8 -ErrorAction Stop | ForEach-Object {
      if (-not [string]::IsNullOrWhiteSpace($_)) {
        try {
          $record = ConvertFrom-Json -InputObject $_ -ErrorAction Stop
          if ($record.SchemaVersion -ne 1 -or -not $record.Kind -or -not $record.Status) { throw 'Invalid history record.' }
          $null = [DateTimeOffset]::Parse($record.StartTime)
          $records.Enqueue($record)
          if ($records.Count -gt $remaining) { $null = $records.Dequeue() }
        } catch { $invalid++ }
      }
    }
    if ($invalid) { Write-Warning ("読み取れない実行履歴を{0}件スキップしました: {1}" -f $invalid, $file.FullName) }
    $latest = $records.ToArray()
    [array]::Reverse($latest)
    $latest
    $remaining -= $latest.Count
    if ($remaining -le 0) { break }
  }
}

Export-ModuleMember -Function Write-WinToolsHistory, Get-WinToolsHistory
