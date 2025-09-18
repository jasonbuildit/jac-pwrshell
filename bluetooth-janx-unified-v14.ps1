# bluetooth-janx-unified-v14.ps1
# PS 7.5 host + Windows PowerShell 5.1 worker (WinRT Bluetooth)
# - Approved verbs for PSScriptAnalyzer
# - PS5.1 helper uses $OpArgs (no $Args collision)
# - WinRT event sniffer with snapshot fallback
# - Back-compat aliases to previous names

param([string]$PreferredAdapterMatch = "5.3")
$script:HasPnpDevice = $false

# ---------------- Approved-verb helpers (PS7) ----------------
function Initialize-BtPnpSupport {
  [CmdletBinding()]
  param()
  try { Import-Module PnpDevice -ErrorAction Stop; $script:HasPnpDevice = $true }
  catch { $script:HasPnpDevice = $false }
}
Initialize-BtPnpSupport

function Get-Ps51Path {
  [CmdletBinding()] param()
  $p = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
  if (-not (Test-Path $p)) { throw "Windows PowerShell 5.1 not found at: $p" }
  $p
}

function Convert-IfJson {
  [CmdletBinding()]
  param($Value, [int]$Depth = 10)
  if ($Value -is [string]) {
    $s = $Value.Trim()
    if ($s.StartsWith('{') -or $s.StartsWith('[')) { return ($s | ConvertFrom-Json -Depth $Depth) }
    $startObj = $s.IndexOf('{'); $startArr = $s.IndexOf('[')
    $starts = @($startObj,$startArr) | Where-Object { $_ -ge 0 } | Sort-Object
    if ($starts.Count -gt 0) {
      $start = $starts[0]; $candidate = $s.Substring($start)
      $endBrace = $candidate.LastIndexOf('}'); $endBracket = $candidate.LastIndexOf(']')
      $end = [Math]::Max($endBrace,$endBracket)
      if ($end -ge 0) { try { return ($candidate.Substring(0,$end+1) | ConvertFrom-Json -Depth $Depth) } catch {} }
    }
    return $s
  }
  $Value
}

# ---------------- Device management (PS7) ----------------
function Get-PairedDevice {
  [CmdletBinding()] param()
  if (-not $script:HasPnpDevice) { Write-Warning "PnpDevice module not available; returning empty."; return @() }
  Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
    Where-Object { $_.FriendlyName } |
    Sort-Object FriendlyName |
    Select-Object Status, FriendlyName, InstanceId
}

function Restart-BtRadio {
  [CmdletBinding()]
  param([string]$Vid='VID_0BDA')
  if (-not $script:HasPnpDevice) { Write-Warning "PnpDevice module not available."; return }
  $d = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
       Where-Object { $_.InstanceId -match $Vid } | Select-Object -First 1
  if (-not $d) { Write-Error "BT radio with $Vid not found."; return }
  Write-Host "Toggling: $($d.FriendlyName) [$($d.InstanceId)]"
  Disable-PnpDevice -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
  Start-Sleep 1.5
  Enable-PnpDevice  -InstanceId $d.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
  Write-Host "Toggled device: $($d.InstanceId)" -ForegroundColor Yellow
}

function Remove-BtDevice {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Match)
  if ($script:HasPnpDevice) {
    $dev = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue | Where-Object {
      $_.InstanceId -eq $Match -or ($_.FriendlyName -and $_.FriendlyName -like "*$Match*")
    } | Select-Object -First 1
    if ($dev) {
      Write-Host "Removing (PnP): $($dev.FriendlyName) [$($dev.InstanceId)]" -ForegroundColor Yellow
      Remove-PnpDevice -InstanceId $dev.InstanceId -Confirm:$false
      Write-Host "Removed." -ForegroundColor Green
      return
    }
  }
  Write-Warning "Falling back to 'pnputil /remove-device'."
  & pnputil.exe /remove-device "$Match" | Write-Host
}

# ---------------- PS 5.1 bridge (WinRT worker) ----------------
function Invoke-PS51 {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Operation,  # Scan, PairId, PairName, Adapters, SetRadio, AdapterInfo, Sniffer
    [hashtable]$OpArgs = @{},
    [switch]$StreamToConsole
  )

  $ps51 = Get-Ps51Path  # this helper should exist already

  # ---------------- PS 5.1 worker script (embedded) ----------------
  $helper = @'
#Requires -Version 5.1
param([Parameter(Mandatory)][string]$Operation,[hashtable]$OpArgs)
$ErrorActionPreference='Stop'
$InformationPreference='SilentlyContinue'
$ProgressPreference='SilentlyContinue'

function Initialize-BtWinRT {
  $tn='Windows.Devices.Enumeration.DeviceInformation, Windows, ContentType=WindowsRuntime'
  $t=[type]::GetType($tn,$false)
  if(-not $t){
    $wm=Join-Path $env:WINDIR 'System32\WinMetadata\Windows.winmd'
    if(Test-Path $wm){ try{ Add-Type -Path $wm -ErrorAction Stop }catch{} }
    $t=[type]::GetType($tn,$false)
  }
  if(-not $t){ throw 'WinRT projection unavailable.' }
}

function Wait-WinRT($op){
  $t=[type]::GetType('Windows.Foundation.AsyncStatus, Windows, ContentType=WindowsRuntime')
  $Started=[enum]::Parse($t,'Started')
  while($op.Status -eq $Started){ Start-Sleep -Milliseconds 50 }
  if($op.PSObject.Methods.Name -contains 'GetResults'){ return $op.GetResults() }
}

function Get-Aqs([ValidateSet('BLE','Classic','All')]$Mode='All'){
  switch($Mode){
    'BLE'     { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}"' }
    'Classic' { 'System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
    default   { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}" OR System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
  }
}

function Find-Devices($Mode='All'){
  Initialize-BtWinRT
  $aqs=Get-Aqs $Mode
  $coll=Wait-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs))
  $coll | ForEach-Object {
    $can=$null;$paired=$null
    try{$can=[bool]$_.Pairing.CanPair}catch{}
    try{$paired=[bool]$_.Pairing.IsPaired}catch{}
    [pscustomobject]@{ Name=$_.Name; Id=$_.Id; CanPair=$can; IsPaired=$paired }
  } | Sort-Object Name
}

function Add-DevicePairingById([string]$Id,[string]$Pin){
  Initialize-BtWinRT
  if ($Id -notlike 'Bluetooth#*') { return @{Status='Error';Error='Provide DeviceInformation.Id starting with Bluetooth#'} }
  $di=Wait-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::CreateFromIdAsync($Id))
  if(-not $di){ return @{Status='Error';Error='DeviceInformation not found'} }
  try{ if($di.Pairing.IsPaired){ return @{Status='AlreadyPaired';Name=$di.Name;Messages=@()} } } catch{}

  $script:PairingPin=$Pin
  $script:OutMsgs=New-Object System.Collections.ArrayList
  $custom=$di.Pairing.Custom
  $h=Register-ObjectEvent -InputObject $custom -EventName PairingRequested -Action {
    $req=$EventArgs
    switch($req.PairingKind){
      'DisplayPin'      { [void]$script:OutMsgs.Add(('PIN: {0}' -f $req.Pin)); $req.Accept() }
      'ConfirmPinMatch' { [void]$script:OutMsgs.Add(('Confirm PIN: {0}' -f $req.Pin)); $req.Accept() }
      'ConfirmOnly'     { [void]$script:OutMsgs.Add('ConfirmOnly'); $req.Accept() }
      'ProvidePin'      { $p=$script:PairingPin; if(-not $p){$p='0000'}; [void]$script:OutMsgs.Add('Providing PIN: ' + $p); $req.Accept($p) }
      default           { [void]$script:OutMsgs.Add('Accept default'); $req.Accept() }
    }
  }

  try{
    $kT=[type]::GetType('Windows.Devices.Enumeration.DevicePairingKinds, Windows, ContentType=WindowsRuntime')
    $k=[enum]::Parse($kT,'ConfirmOnly') -bor [enum]::Parse($kT,'DisplayPin') -bor [enum]::Parse($kT,'ConfirmPinMatch') -bor [enum]::Parse($kT,'ProvidePin')
    $lT=[type]::GetType('Windows.Devices.Enumeration.DevicePairingProtectionLevel, Windows, ContentType=WindowsRuntime')
    $Def=[enum]::Parse($lT,'Default'); $None=[enum]::Parse($lT,'None')

    $r =Wait-WinRT ($custom.PairAsync($Def,$k))
    if($r.Status  -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired){
      return @{Status='Paired';Name=$di.Name;Messages=$script:OutMsgs}
    }

    $r2=Wait-WinRT ($di.Pairing.PairAsync($None))
    if($r2.Status -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired){
      return @{Status='PairedBasic';Name=$di.Name;Messages=$script:OutMsgs}
    }

    $r3=Wait-WinRT ($di.Pairing.PairAsync($Def))
    if($r3.Status -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired){
      return @{Status='PairedDefault';Name=$di.Name;Messages=$script:OutMsgs}
    }

    return @{Status='Failed';Name=$di.Name;Details="$($r.Status)/$($r2.Status)/$($r3.Status)";Messages=$script:OutMsgs}
  }
  finally {
    if($h){ Unregister-Event -SourceIdentifier $h.Name | Out-Null }
  }
}

function Add-DevicePairingByName([string]$Name,[int]$Retries=3,[int]$Between=2,[string]$Pin){
  $rx=New-Object System.Text.RegularExpressions.Regex ([regex]::Escape($Name)),'IgnoreCase'
  for($i=1;$i -le $Retries;$i++){
    $d=Find-Devices All | Where-Object { $_.Name -and $rx.IsMatch($_.Name) -and (-not $_.IsPaired) } | Select-Object -First 1
    if($d){ return (Add-DevicePairingById -Id $d.Id -Pin $Pin) }
    Write-Host ("[{0}/{1}] No discoverable match for '{2}'. Re-scanning..." -f $i,$Retries,$Name)
    Start-Sleep -Seconds $Between
  }
  return @{Status='NoMatch';Name=$Name}
}

function Get-BtAdapters51 {
  Initialize-BtWinRT
  $radios=Wait-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
  $bt=$radios | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }
  $bt | ForEach-Object { [pscustomobject]@{ Name=$_.Name; State=$_.State; Id=$_.DeviceId } } | Sort-Object Name
}

function Set-BtRadioState51([string]$NameOrId,[string]$State){
  Initialize-BtWinRT
  $radios=Wait-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
  $bt=$radios | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }
  $m=$bt | Where-Object { $_.DeviceId -eq $NameOrId -or ($_.Name -and $_.Name -like "*$NameOrId*") } | Select-Object -First 1
  if(-not $m){ return @{ Status='Error'; Error="No adapter matched '$NameOrId'." } }
  $target= if($State -eq 'On'){ [Windows.Devices.Radios.RadioState]::On } else { [Windows.Devices.Radios.RadioState]::Off }
  [void](Wait-WinRT ($m.SetStateAsync($target)))
  return @{ Name=$m.Name; State="$($target)" }
}

function Get-BtAdapterInfo51 {
  Initialize-BtWinRT
  $a=$null; try{$a=Wait-WinRT ([Windows.Devices.Bluetooth.BluetoothAdapter]::GetDefaultAsync())}catch{}
  $r=$null; if($a){try{$r=Wait-WinRT ($a.GetRadioAsync())}catch{}}
  if(-not $r){
    $rs=Wait-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
    $r=$rs | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth } | Select-Object -First 1
  }
  $addr=$null; if($a -and $a.BluetoothAddress){ $addr=('{0:X12}' -f $a.BluetoothAddress) }
  [pscustomobject]@{
    HasDefaultAdapter=[bool]$a
    BluetoothAddressHex=$addr
    IsLowEnergySupported=($a -and $a.IsLowEnergySupported)
    IsClassicSupported=($a -and $a.IsClassicSupported)
    RadioName=($r.Name)
    RadioState=($r.State)
    RadioId=($r.DeviceId)
  }
}

function Start-BtSniffer51([int]$Seconds=15,[string]$NameLike,[string]$Mode='All'){
  Initialize-BtWinRT
  $filterRegex = $null
  if ($NameLike) { $filterRegex = New-Object System.Text.RegularExpressions.Regex ($NameLike),'IgnoreCase' }

  $watcher = $null; $eventId='ADV'; $fastOk=$false
  try {
    $watcher = [Windows.Devices.Bluetooth.Advertisement.BluetoothLEAdvertisementWatcher]::new()
    $scanModeType = [type]::GetType('Windows.Devices.Bluetooth.Advertisement.BluetoothLEScanningMode, Windows, ContentType=WindowsRuntime')
    if ($scanModeType) { $Active=[enum]::Parse($scanModeType,'Active'); try{ $watcher.ScanningMode=$Active }catch{} }

    Register-ObjectEvent -InputObject $watcher -EventName Received -SourceIdentifier $eventId -Action {
      try{
        $e=$EventArgs; $name=$e.Advertisement.LocalName
        if ($script:__flt -and -not $script:__flt.IsMatch([string]$name)) { return }
        $mac=('{0:X12}' -f $e.BluetoothAddress) -replace '(.{2})(?=.)','$1:'
        $svc=$null; try{ $uu=@(); foreach($u in $e.Advertisement.ServiceUuids){ $uu+= $u.ToString() }; if($uu){ $svc=$uu -join ',' } }catch{}
        Write-Host ("ADV  MAC={0,-17} RSSI={1,4}dBm  Name={2}  Services=[{3}]" -f $mac,$e.RawSignalStrengthInDBm,($name -as [string]),($svc -as [string])) -ForegroundColor Gray
      }catch{}
    } | Out-Null

    if ($filterRegex) { $script:__flt=$filterRegex } else { $script:__flt=$null }
    $watcher.Start(); Start-Sleep -Seconds $Seconds; $watcher.Stop()
    $fastOk=$true
  } catch {
    # fall back to snapshot
  } finally {
    if ($fastOk -and $watcher) { Unregister-Event -SourceIdentifier $eventId -ErrorAction SilentlyContinue | Out-Null }
    Remove-Variable -Name __flt -Scope Script -ErrorAction SilentlyContinue
  }

  if (-not $fastOk) {
    Write-Warning "WinRT 'Received' event not available; using snapshot fallback (no RSSI)."
    $aqs=Get-Aqs $Mode
    $tEnd=(Get-Date).AddSeconds($Seconds)
    while((Get-Date) -lt $tEnd){
      try{
        $coll=[Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs)
        while($coll.Status -eq ([Windows.Foundation.AsyncStatus]::Started)){ Start-Sleep -Milliseconds 50 }
        $list=$coll.GetResults()
        $rows=foreach($di in $list){
          $nm=$di.Name; if($filterRegex -and -not $filterRegex.IsMatch([string]$nm)){ continue }
          $can=$null;$paired=$null; try{$can=[bool]$di.Pairing.CanPair}catch{}; try{$paired=[bool]$di.Pairing.IsPaired}catch{}
          [pscustomobject]@{ Name=$nm; CanPair=$can; IsPaired=$paired; Id=$di.Id }
        }
        if($rows){
          Write-Host ("--- Snapshot {0:HH:mm:ss} (Mode={1}) ---" -f (Get-Date), $Mode) -ForegroundColor DarkCyan
          $rows | Sort-Object Name | Format-Table Name,CanPair,IsPaired -AutoSize
        }
      }catch{}
      Start-Sleep -Milliseconds 800
    }
  }
}

# Frame JSON outputs so the caller can reliably parse
$__json = $null
switch($Operation){
  'Scan'        { $__json = (Find-Devices ($OpArgs.Mode) | ConvertTo-Json -Depth 6); break }
  'PairId'      { $__json = ((Add-DevicePairingById -Id $OpArgs.Id -Pin $OpArgs.Pin) | ConvertTo-Json -Depth 6); break }
  'PairName'    { $__json = ((Add-DevicePairingByName -Name $OpArgs.Name -Retries $OpArgs.Retries -Between $OpArgs.Between -Pin $OpArgs.Pin) | ConvertTo-Json -Depth 6); break }
  'Adapters'    { $__json = (Get-BtAdapters51 | ConvertTo-Json -Depth 5); break }
  'SetRadio'    { $__json = ((Set-BtRadioState51 -NameOrId $OpArgs.NameOrId -State $OpArgs.State) | ConvertTo-Json -Depth 4); break }
  'AdapterInfo' { $__json = (Get-BtAdapterInfo51 | ConvertTo-Json -Depth 6); break }
  'Sniffer'     {
    $modeArg='All'
    if ($OpArgs -and ($OpArgs.ContainsKey('Mode')) -and $OpArgs.Mode) { $modeArg = $OpArgs.Mode }
    Start-BtSniffer51 -Seconds ([int]$OpArgs.Seconds) -NameLike $OpArgs.NameLike -Mode $modeArg
    break
  }
  default       { $__json = (@{ Status='Error'; Error=("Unknown operation: {0}" -f $Operation) } | ConvertTo-Json -Depth 3); break }
}

if ($__json) {
  # Emit framed JSON only
  Write-Output '<<<JSON_START>>>' 
  Write-Output $__json
  Write-Output '<<<JSON_END>>>'
}
'@

  # ----- payload from PS7 -----
  $payload = @{ Operation = $Operation; OpArgs = $OpArgs } | ConvertTo-Json -Depth 6

  # Convert PSCustomObject OpArgs -> Hashtable for PS5.1 param binding
  $runner = @"
`$ErrorActionPreference='Stop'
`$input = @'
$payload
'@ | ConvertFrom-Json

`$ht = @{}
if (`$input.OpArgs) {
  foreach (`$p in `$input.OpArgs.PSObject.Properties) { `$ht[`$p.Name] = `$p.Value }
}

& {
$helper
} -Operation `$input.Operation -OpArgs `$ht
"@

  # Run
  $raw = & $ps51 -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $runner 2>&1 | Out-String

  if ($StreamToConsole) {
    $raw | Write-Host
    return
  }

  # Extract framed JSON (if present); else return raw text
  if ($raw -match '<<<JSON_START>>>(?<json>.*)<<<JSON_END>>>'s) {
    $json = $Matches['json'].Trim()
    try { return ($json | ConvertFrom-Json -Depth 10) } catch { return $json }
  } else {
    return $raw
  }
}

# ---------------- Public wrappers (approved verbs) ----------------
function Get-BtAdapters { Convert-IfJson (Invoke-PS51 -Operation Adapters) | Sort-Object Name }

function Get-BtAdapterInfo { Convert-IfJson (Invoke-PS51 -Operation AdapterInfo) }

function Set-BtAdapterState {
  [CmdletBinding()]
  param([Parameter(Mandatory)][ValidateSet('On','Off')]$State,[Parameter(Mandatory)][string]$NameOrId)
  [void](Invoke-PS51 -Operation SetRadio -OpArgs @{ NameOrId=$NameOrId; State=$State })
  Get-BtAdapters | Where-Object { $_.Name -like "*$NameOrId*" -or $_.Id -eq $NameOrId }
}

function Find-BtDevice {
  [CmdletBinding()]
  param([ValidateSet('BLE','Classic','All')]$Mode='All')
  Convert-IfJson (Invoke-PS51 -Operation Scan -OpArgs @{ Mode=$Mode }) | Sort-Object Name
}

function Add-BtDevicePairing {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$DeviceIdOrAddress,[string]$Pin)
  if ($DeviceIdOrAddress -notlike 'Bluetooth#*') { throw "Provide DeviceInformation.Id starting with 'Bluetooth#'." }
  Convert-IfJson (Invoke-PS51 -Operation PairId -OpArgs @{ Id=$DeviceIdOrAddress; Pin=$Pin })
}

function Add-BtDevicePairingByName {
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$NameMatch,[string]$Pin,[int]$Retries=8,[int]$BetweenSeconds=2)
  Convert-IfJson (Invoke-PS51 -Operation PairName -OpArgs @{ Name=$NameMatch; Retries=$Retries; Between=$BetweenSeconds; Pin=$Pin })
}

function Start-BtSniffer {
  [CmdletBinding()]
  param([int]$Seconds=15,[string]$NameLike,[ValidateSet('BLE','Classic','All')]$Mode='All')
  Write-Host "Starting live sniffer (PS 5.1 child)..." -ForegroundColor Cyan
  Invoke-PS51 -Operation Sniffer -OpArgs @{ Seconds=$Seconds; NameLike=$NameLike; Mode=$Mode } -StreamToConsole
}

function Start-BtSnifferWindow {
  [CmdletBinding()]
  param([int]$Seconds = 15, [string]$NameLike,[ValidateSet('BLE','Classic','All')]$Mode='All')
  # For simplicity reuse the same worker pipeline but in new window:
  Start-Process (Get-Ps51Path) -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-Command',
    ("`$s={0};`$f='{1}';`$m='{2}';" -f $Seconds, ($NameLike -replace "'", "''"), $Mode) + "
    `"`$(Get-Content -Raw -LiteralPath '$PSCommandPath')`";
    Invoke-PS51 -Operation Sniffer -OpArgs @{ Seconds=`$s; NameLike=`$f; Mode=`$m } -StreamToConsole"
  )
}

function Start-BtKeyboardListener {
  [CmdletBinding()]
  param(
    [string]$NameMatch = 'Keyboard K380',
    [ValidateSet('BLE','Classic','All')] [string]$Mode = 'All',
    [int]$TimeoutMinutes = 5,
    [string]$Pin
  )
  $seconds = [int]([TimeSpan]::FromMinutes($TimeoutMinutes).TotalSeconds)
  Start-BtSnifferWindow -Seconds $seconds -NameLike $NameMatch -Mode $Mode | Out-Null
  Write-Host "Listening for '$NameMatch' and attempting auto-pair… (Mode=$Mode, Timeout=$TimeoutMinutes min)" -ForegroundColor Cyan
  $stopAt = (Get-Date).AddSeconds($seconds)
  $paired = $false
  do {
    $result = Add-BtDevicePairingByName -NameMatch $NameMatch -Retries 1 -BetweenSeconds 1 -Pin $Pin
    if ($result -and ($result.Status -match '^Paired')) {
      Write-Host ("Paired: {0}  [{1}]" -f $result.Name, $result.Status) -ForegroundColor Green
      $paired = $true
      break
    }
    Start-Sleep -Seconds 1
  } while ((Get-Date) -lt $stopAt)
  if (-not $paired) { Write-Warning "Finished listening window without a successful pair for '$NameMatch'." }
}

function Start-BtAnyDeviceListener {
  [CmdletBinding()]
  param(
    [int]$TimeoutMinutes = 3,
    [int]$ScanSnapshotEverySeconds = 10,
    [ValidateSet('BLE','Classic','All')] [string]$Mode = 'All'
  )
  $seconds = [int]([TimeSpan]::FromMinutes($TimeoutMinutes).TotalSeconds)
  Start-BtSnifferWindow -Seconds $seconds -NameLike '' -Mode $Mode | Out-Null
  Write-Host "Listening for ANY Bluetooth advertising for $TimeoutMinutes min… (live ADV in separate window)" -ForegroundColor Cyan
  $stopAt = (Get-Date).AddSeconds($seconds)
  if ($ScanSnapshotEverySeconds -gt 0) {
    while ((Get-Date) -lt $stopAt) {
      try {
        $snap = Find-BtDevice -Mode $Mode
        if ($snap) {
          Write-Host ("--- Snapshot ({0:HH:mm:ss}) discoverable devices ---" -f (Get-Date)) -ForegroundColor DarkCyan
          $snap | Sort-Object Name | Format-Table Name, CanPair, IsPaired -AutoSize
        }
      } catch { Write-Warning $_.Exception.Message }
      $remain = [int]([TimeSpan]::FromTicks(($stopAt - (Get-Date)).Ticks).TotalSeconds)
      if ($remain -le 0) { break }
      Start-Sleep -Seconds ([Math]::Min($ScanSnapshotEverySeconds, $remain))
    }
  } else {
    Start-Sleep -Seconds $seconds
  }
  Write-Host "Listener finished." -ForegroundColor Yellow
}

# ---------------- Back-compat aliases ----------------
Set-Alias Ensure-PnpDeviceModule Initialize-BtPnpSupport -ErrorAction SilentlyContinue
Set-Alias Toggle-BtRadio       Restart-BtRadio         -ErrorAction SilentlyContinue
Set-Alias Find-BtDevices       Find-BtDevice           -ErrorAction SilentlyContinue
Set-Alias Pair-BtDevice        Add-BtDevicePairing     -ErrorAction SilentlyContinue
Set-Alias Pair-BtByName        Add-BtDevicePairingByName -ErrorAction SilentlyContinue
Set-Alias Start-BleSniffer     Start-BtSniffer         -ErrorAction SilentlyContinue
Set-Alias Start-Ps51SnifferWindow Start-BtSnifferWindow -ErrorAction SilentlyContinue
Set-Alias Listen-ForKeyboard   Start-BtKeyboardListener -ErrorAction SilentlyContinue
Set-Alias Listen-ForAnyDevice  Start-BtAnyDeviceListener -ErrorAction SilentlyContinue
Set-Alias Get-PairedDevices    Get-PairedDevice        -ErrorAction SilentlyContinue

# ---------------- Help banner ----------------
Write-Host "Loaded bluetooth-janx-unified-v14.ps1 (PS7 bridge + 5.1 WinRT, approved verbs)." -ForegroundColor Cyan
@"
Get-BtAdapterInfo
Get-BtAdapters
Set-BtAdapterState -State On -NameOrId '5.3'
Find-BtDevice -Mode All | ft
Add-BtDevicePairingByName -NameMatch 'Keyboard K380' -Retries 10 -BetweenSeconds 1 -Pin '777036'
Add-BtDevicePairing -DeviceIdOrAddress 'Bluetooth#Bluetooth...'
Start-BtSniffer -Seconds 12 -NameLike 'Logi'            # inline; falls back if WinRT events blocked
Start-BtSnifferWindow -Seconds 20 -NameLike ''
Start-BtKeyboardListener -NameMatch 'Keyboard K380' -TimeoutMinutes 3 -Pin '777036'
Start-BtAnyDeviceListener -TimeoutMinutes 3 -ScanSnapshotEverySeconds 10 -Mode All
Get-PairedDevice
Restart-BtRadio
Remove-BtDevice -Match 'Logitech'
"@ | Write-Host
