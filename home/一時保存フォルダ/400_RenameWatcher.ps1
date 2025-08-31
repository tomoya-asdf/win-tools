# @Powershell -NoP -W Hidden -C "$PSCP='%~f0';$PSSR='%~dp0'.TrimEnd('\');&([ScriptBlock]::Create((gc '%~f0'|?{$_.ReadCount -gt 1}|Out-String)))" %* & exit/b
# by earthdiver1  V1.05
if ($PSCommandPath) {
    $PSCP = $PSCommandPath
    $PSSR = $PSScriptRoot
    $code = '[DllImport("user32.dll")]public static extern bool ShowWindowAsync(IntPtr hWnd,int nCmdShow);'
    $type = Add-Type -MemberDefinition $code -Name Win32ShowWindowAsync -PassThru
    [void]$type::ShowWindowAsync((Get-Process -PID $PID).MainWindowHandle,0) }
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$menuItem = New-Object System.Windows.Forms.MenuItem "Exit"
$menuItem.add_Click({$notifyIcon.Visible=$False;while(-not $status.IsCompleted){Start-Sleep 1};$appContext.ExitThread()})
$contextMenu = New-Object System.Windows.Forms.ContextMenu
$contextMenu.MenuItems.AddRange($menuItem)
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.ContextMenu = $contextMenu
$notifyIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSCP)
$notifyIcon.Text = (Get-ChildItem $PSCP).BaseName
$notifyIcon.Visible = $True
$_syncHash = [hashtable]::Synchronized(@{})
$_syncHash.NI   = $notifyIcon
$_syncHash.PSCP = $PSCP
$_syncHash.PSSR = $PSSR
$runspace = [RunspaceFactory]::CreateRunspace()
$runspace.ApartmentState = "STA"
$runspace.ThreadOptions  = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("_syncHash",$_syncHash)
$scriptBlock = Get-Content $PSCP | ?{ $on -or $_[1] -eq "!" }| %{ $on=1; $_ } | Out-String
$action=[ScriptBlock]::Create(@'
#   param($Param1, $Param2)
    Start-Transcript -LiteralPath ($_syncHash.PSCP -Replace '\..*?$',".log") -Append
    Function Start-Sleep { [CmdletBinding(DefaultParameterSetName="S")]
        param([parameter(Position=0,ParameterSetName="M")][Int]$Milliseconds,
              [parameter(Position=0,ParameterSetName="S")][Int]$Seconds,[Switch]$NoExit)
        if ($PsCmdlet.ParameterSetName -eq "S") {
            $int = 5
            for ($i = 0; $i -lt $Seconds; $i += $int) {
                if (-not($NoExit -or $_syncHash.NI.Visible)) { exit }
                Microsoft.PowerShell.Utility\Start-Sleep -Seconds $int }
        } else {
            $int = 100
            for ($i = 0; $i -lt $Milliseconds; $i += $int) {
                if (-not($NoExit -or $_syncHash.NI.Visible)) { exit }
                Microsoft.PowerShell.Utility\Start-Sleep -Milliseconds $int }}}
    $script:PSCommandPath = $_syncHash.PSCP
    $script:PSScriptRoot  = $_syncHash.PSSR
'@ + $scriptBlock)
$PS = [PowerShell]::Create().AddScript($action) #.AddArgument($Param1).AddArgument($Param2)
$PS.Runspace = $runspace
$status = $PS.BeginInvoke()
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)
exit
# ! ---------- ScriptBlock (Line No. 28) begins here ---------- DO NOT REMOVE THIS LINE

# スクリプト自身のあるフォルダを取得 
$scriptPath = $MyInvocation.MyCommand.Path 
$watchPath = Split-Path -Path $scriptPath -Parent

# FileSystemWatcher の設定
$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchPath
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, DirectoryName'

# イベントハンドラの登録
$createdHandler = Register-ObjectEvent $watcher "Created" -Action {
    $itemPath = $Event.SourceEventArgs.FullPath
    $itemName = $Event.SourceEventArgs.Name

    Write-Host "Createイベントを検知しました。: $itemPath"
    # このファイルを除外
    if ($itemPath -eq $scriptPath) { return }
    # ~で始まる名称を除外
    if ($itemName -match '^~') { return }
    # [数字3桁]+_で始まる名称を除外
    if ($itemName -match '^\d{3}_') { return }
    # 【一時保存】で始まる名称を除外
    if ($itemName -match '^【一時保存】') { return }

    # 名称の変更
    $dateTag = Get-Date -Format "yyyyMMdd"
    $newName = "【一時保存】＜" + $dateTag + "＞" + $itemName
    $newPath = Join-Path -Path (Split-Path $itemPath) -ChildPath $newName

    Start-Sleep -Seconds 1
    try {
        Rename-Item -Path $itemPath -NewName $newName
        Write-Host "Renamed: $itemName → $newName"
    } catch {
        Write-Host "Failed to rename [$itemName] $_"
    }
 }

Write-Host "監視を開始します: $watchPath"
Write-Host "Enterキーで終了します..."

# 無限ループで監視を維持
while ($true) {
    Start-Sleep -Seconds 1
    if ([Console]::KeyAvailable) {
        $key = [Console]::ReadKey($true).Key
        if ($key -eq "Enter") {
            break
        }
    }
}

# イベント解除
Unregister-Event -SourceIdentifier $createdHandler.Name
