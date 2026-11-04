$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$eventScript = Join-Path $repoRoot 'tools\logs\Get-WorkEventLog.ps1'
if (-not ('WinTools.UsnJournal' -as [type])) {
  Add-Type -Path (Join-Path $repoRoot 'tools\logs\lib\UsnJournal.cs') -ErrorAction Stop
}

function New-TestUsnBuffer {
  param(
    [string]$Name = '案件フォルダ',
    [uint32]$Reason = 0x100,
    [uint32]$Attributes = 0x10,
    [long]$Usn = 100,
    [uint16]$Version = 2
  )
  $nameBytes = [Text.Encoding]::Unicode.GetBytes($Name)
  $size = [int]([Math]::Ceiling((60 + $nameBytes.Length) / 8.0) * 8)
  $buffer = New-Object byte[] (8 + $size)
  [Array]::Copy([BitConverter]::GetBytes([long]($Usn + $size)), 0, $buffer, 0, 8)
  [Array]::Copy([BitConverter]::GetBytes([uint32]$size), 0, $buffer, 8, 4)
  [Array]::Copy([BitConverter]::GetBytes($Version), 0, $buffer, 12, 2)
  [Array]::Copy([BitConverter]::GetBytes([uint64]123), 0, $buffer, 16, 8)
  [Array]::Copy([BitConverter]::GetBytes([uint64]456), 0, $buffer, 24, 8)
  [Array]::Copy([BitConverter]::GetBytes($Usn), 0, $buffer, 32, 8)
  [Array]::Copy([BitConverter]::GetBytes(([datetime]'2026-11-01T00:00:00Z').ToFileTimeUtc()), 0, $buffer, 40, 8)
  [Array]::Copy([BitConverter]::GetBytes($Reason), 0, $buffer, 48, 4)
  [Array]::Copy([BitConverter]::GetBytes($Attributes), 0, $buffer, 60, 4)
  [Array]::Copy([BitConverter]::GetBytes([uint16]$nameBytes.Length), 0, $buffer, 64, 2)
  [Array]::Copy([BitConverter]::GetBytes([uint16]60), 0, $buffer, 66, 2)
  [Array]::Copy($nameBytes, 0, $buffer, 68, $nameBytes.Length)
  return ,$buffer
}

function New-TestEvent {
  param([string]$Log = 'System', [long]$Record = 1, [datetime]$Time = '2026-11-01 10:00', [int]$Id = 41)
  $event = [pscustomobject]@{
    TimeCreated = $Time; LogName = $Log; RecordId = $Record; Id = $Id
    ProviderName = 'TestProvider'; LevelDisplayName = 'Error'; MachineName = 'TEST-PC'; Message = 'Test message'
  }
  $event | Add-Member ScriptMethod ToXml {
    '<Event xmlns="http://schemas.microsoft.com/win/2004/08/events/event"><EventData><Data Name="TargetUserName">test-user</Data><Data Name="TargetDomainName">TEST</Data><Data Name="LogonType">10</Data><Data Name="IpAddress">192.0.2.1</Data></EventData></Event>'
  }
  return $event
}

Describe 'NTFS USN binary parsing' {
  It 'reads Japanese folder names, IDs and local timestamps' {
    $buffer = New-TestUsnBuffer
    $rows = [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000)
    $rows.Count | Should Be 1
    $rows[0].Name | Should Be '案件フォルダ'
    $rows[0].FileId | Should Be 123
    $rows[0].ParentFileId | Should Be 456
    $rows[0].TimeCreated.ToUniversalTime() | Should Be ([datetime]'2026-11-01T00:00:00Z').ToUniversalTime()
    $rows[0].PathHint | Should BeNullOrEmpty
  }
  It 'keeps deleted directories without checking their existence' {
    $buffer = New-TestUsnBuffer -Reason 0x200
    [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000)[0].Operation | Should Be 'Deleted'
  }
  It 'excludes files and close-only records' {
    $buffer = New-TestUsnBuffer -Attributes 0x20
    [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000).Count | Should Be 0
    $buffer = New-TestUsnBuffer -Reason ([uint32]2147483648)
    [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000).Count | Should Be 0
  }
  It 'preserves both rename reasons when aggregated' {
    $buffer = New-TestUsnBuffer -Reason 0x3000
    [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000)[0].Operation | Should Be 'RenamedFrom, RenamedTo'
  }
  It 'excludes changes at or after the snapshot tail' {
    $buffer = New-TestUsnBuffer -Usn 100
    [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 100).Count | Should Be 0
  }
  It 'walks variable-length records using RecordLength' {
    $first = New-TestUsnBuffer -Name '短' -Usn 100
    $second = New-TestUsnBuffer -Name '長いフォルダ名' -Usn 200
    $combined = New-Object byte[] ($first.Length + $second.Length - 8)
    [Array]::Copy($first, $combined, $first.Length)
    [Array]::Copy($second, 8, $combined, $first.Length, $second.Length - 8)
    $rows = [WinTools.UsnJournal]::ParseBuffer($combined, $combined.Length, 1000)
    $rows.Count | Should Be 2
    $rows[1].Name | Should Be '長いフォルダ名'
  }
  It 'rejects truncated buffers, invalid lengths and invalid name bounds' {
    { [WinTools.UsnJournal]::ParseBuffer((New-Object byte[] 4), 4, 1000) } | Should Throw
    $buffer = New-TestUsnBuffer
    $buffer[8] = 0; $buffer[9] = 0
    { [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000) } | Should Throw
    $buffer = New-TestUsnBuffer
    $buffer[64] = 255
    { [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000) } | Should Throw
  }
  It 'fails explicitly for unknown record versions' {
    $buffer = New-TestUsnBuffer -Version 3
    { [WinTools.UsnJournal]::ParseBuffer($buffer, $buffer.Length, 1000) } | Should Throw
  }
}

Describe 'Work event log queries' {
  It 'sorts across logs before applying the global limit' {
    Mock Get-WinEvent {
      if ($FilterHashtable.LogName -eq 'System') {
        New-TestEvent -Record 1 -Time '2026-11-01 10:00'
      } else {
        New-TestEvent -Log Application -Record 2 -Time '2026-11-01 11:00'
      }
    }
    $rows = @(& $eventScript -MaxEvents 1)
    $rows.Count | Should Be 1
    $rows[0].LogName | Should Be 'Application'
    Assert-MockCalled Get-WinEvent -Times 2 -Exactly -Scope It -ParameterFilter { $MaxEvents -eq 1 -and $FilterHashtable.Level.Count -eq 3 }
  }
  It 'merges an event matching Errors and Restart without losing labels' {
    Mock Get-WinEvent { New-TestEvent }
    $rows = @(& $eventScript -Category Errors,Restart)
    $rows.Count | Should Be 1
    $rows[0].Category | Should Be 'Errors, Restart'
  }
  It 'reads logon data by XML name rather than localized message positions' {
    Mock Get-WinEvent { New-TestEvent -Log Security -Id 4624 }
    $rows = @(& $eventScript -Category Logon)
    $rows[0].User | Should Be 'TEST\test-user'
    $rows[0].LogonType | Should Be '10'
    $rows[0].IpAddress | Should Be '192.0.2.1'
    Assert-MockCalled Get-WinEvent -Times 1 -Exactly -Scope It -ParameterFilter {
      $FilterHashtable.LogName -eq 'Security' -and
      $FilterHashtable.ProviderName -eq 'Microsoft-Windows-Security-Auditing' -and
      $FilterHashtable.Id -contains 4801
    }
  }
  It 'continues other logs when one log denies access' {
    Mock Get-WinEvent {
      if ($FilterHashtable.LogName -eq 'Security') { throw 'Access denied' }
      New-TestEvent
    }
    $warnings = @()
    $rows = @(& $eventScript -Category Logon,Errors -WarningVariable warnings -WarningAction SilentlyContinue)
    $rows.Count | Should Be 1
    $warnings.Count | Should Be 1
    $warnings[0].ToString() | Should Match 'Security'
  }
  It 'treats no matching events as an empty result without a permission warning' {
    Mock Get-WinEvent { Write-Error 'No records' -ErrorId NoMatchingEventsFound -ErrorAction Stop }
    $warnings = @()
    $rows = @(& $eventScript -WarningVariable warnings -WarningAction SilentlyContinue)
    $rows.Count | Should Be 0
    $warnings.Count | Should Be 0
  }
  It 'validates time ranges before querying logs' {
    Mock Get-WinEvent { throw 'Should not be called' }
    { & $eventScript -StartTime '2026-11-02' -EndTime '2026-11-01' } | Should Throw
    Assert-MockCalled Get-WinEvent -Times 0 -Exactly -Scope It
  }
  It 'preserves the event when supplemental XML cannot be read' {
    Mock Get-WinEvent {
      $event = New-TestEvent
      $event | Add-Member ScriptMethod ToXml { throw 'Missing XML' } -Force
      $event
    }
    $rows = @(& $eventScript -Category Errors)
    $rows.Count | Should Be 1
    $rows[0].Id | Should Be 41
    $rows[0].User | Should BeNullOrEmpty
  }
  It 'passes the requested date bounds to the event log service' {
    Mock Get-WinEvent { New-TestEvent }
    & $eventScript -Category Logon -StartTime '2026-11-01 09:00' -EndTime '2026-11-01 18:00' | Out-Null
    Assert-MockCalled Get-WinEvent -Times 1 -Exactly -Scope It -ParameterFilter {
      $FilterHashtable.StartTime -eq [datetime]'2026-11-01 09:00' -and
      $FilterHashtable.EndTime -eq [datetime]'2026-11-01 18:00'
    }
  }
  It 'keeps distinct events with the same record ID in different logs' {
    Mock Get-WinEvent { New-TestEvent -Log $FilterHashtable.LogName }
    @(& $eventScript).Count | Should Be 2
  }
  It 'does not query categories twice when All or duplicate names are supplied' {
    Mock Get-WinEvent { New-TestEvent -Log $FilterHashtable.LogName }
    & $eventScript -Category All,Errors,All | Out-Null
    Assert-MockCalled Get-WinEvent -Times 8 -Exactly -Scope It
  }
}
