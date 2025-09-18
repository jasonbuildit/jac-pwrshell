# bluetooth-janx-ps7.ps1  — PowerShell 7.5 shim for your Bluetooth tools
# Run this in pwsh 7.5 on Windows. Uses a Windows PowerShell 5.1 child for WinRT calls.

param(
  [string]$PreferredAdapterMatch = "5.3"
)

# -------------------- Utilities: PS5.1 bridge --------------------
function Get-Ps51Path {
  $p = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
  if (-not (Test-Path $p)) { throw "Windows PowerShell 5.1 not found at: $p" }
  return $p
}

# Invoke a small PS5.1 helper script with arguments and capture JSON / text
function Invoke-PS51 {
  param(
    [Parameter(Mandatory)][string]$Operation,  # Scan, PairId, PairName, Sniffer, ListenKb, ListenAny, Adapters, SetRadio, AdapterInfo
    [hashtable]$Args = @{},
    [switch]$StreamToConsole  # for live feeds (sniffer/listeners)
  )
  $ps51 = Get-Ps51Path

  # --- Embedded PS5.1 helper (WinRT + operations) ---
  $helper = @'
#Requires -Version 5.1
param([Parameter(Mandatory)][string]$Operation,[hashtable]$Args)

function Ensure-WinRT {
  $tn = 'Windows.Devices.Enumeration.DeviceInformation, Windows, ContentType=WindowsRuntime'
  $t  = [type]::GetType($tn, $false)
  if (-not $t) {
    $winmd = Join-Path $env:WINDIR 'System32\WinMetadata\Windows.winmd'
    if (Test-Path $winmd) { try { Add-Type -Path $winmd -ErrorAction Stop } catch { } }
    $t = [type]::GetType($tn, $false)
  }
  if (-not $t) { throw "WinRT projection unavailable." }
}
function Await-WinRT([Parameter(Mandatory)]$op) {
  $AsyncStatusType = [type]::GetType('Windows.Foundation.AsyncStatus, Windows, ContentType=WindowsRuntime')
  $Started  = [enum]::Parse($AsyncStatusType,'Started')
  while ($op.Status -eq $Started) { Start-Sleep -Milliseconds 50 }
  if ($op.PSObject.Methods.Name -contains 'GetResults') { return $op.GetResults() }
  return $null
}
function Get-Aqs([ValidateSet("BLE","Classic","All")]$Mode="All") {
  switch($Mode){
    'BLE'     {'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}"'}
    'Classic' {'System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"'}
    default   {'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}" OR System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"'}
  }
}
function Find-Devices([string]$Mode='All){
  Ensure-WinRT
  $aqs  = Get-Aqs $Mode
  $coll = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs))
  $items = foreach($di in $coll){
    [pscustomobject]@{ Name=$di.Name; Id=$di.Id; IsPaired=(try{[bool]$di.Pairing.IsPaired}catch{$null}) }
  }
  $items | Sort-Object Name
}
function Pair-ById([string]$Id,[string]$Pin){
  Ensure-WinRT
  if ($Id -notlike "Bluetooth#*") { throw "Provide DeviceInformation.Id starting with 'Bluetooth#'." }
  $di = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::CreateFromIdAsync($Id))
  if (-not $di){ throw "DeviceInformation not found." }
  try{ if($di.Pairing.IsPaired){ return @{ Status="AlreadyPaired"; Name=$di.Name } } }catch{}
  $script:PairingPin=$Pin
  $custom=$di.Pairing.Custom
  $handler = Register-ObjectEvent -InputObject $custom -EventName PairingRequested -Action {
    $req=$EventArgs
    switch($req.PairingKind){
      "DisplayPin"      { Write-Host ("PIN: {0}" -f $req.Pin); $req.Accept() }
      "ConfirmPinMatch" { Write-Host ("Confirm PIN: {0}" -f $req.Pin); $req.Accept() }
      "ConfirmOnly"     { $req.Accept() }
      "ProvidePin"      { $p=$script:PairingPin; if(-not $p){$p="0000"}; $req.Accept($p) }
      default           { $req.Accept() }
    }
  }
  try{
    $kindsType=[type]::GetType('Windows.Devices.Enumeration.DevicePairingKinds, Windows, ContentType=WindowsRuntime')
    $kinds = [enum]::Parse($kindsType,'ConfirmOnly') -bor [enum]::Parse($kindsType,'DisplayPin') -bor [enum]::Parse($kindsType,'ConfirmPinMatch') -bor [enum]::Parse($kindsType,'ProvidePin')
    $levelType=[type]::GetType('Windows.Devices.Enumeration.DevicePairingProtectionLevel, Windows, ContentType=WindowsRuntime')
    $Default=[enum]::Parse($levelType,'Default'); $None=[enum]::Parse($levelType,'None')
    $res = Await-WinRT ($custom.PairAsync($Default,$kinds))
    if ($res.Status -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired) { return @{ Status="Paired"; Name=$di.Name } }
    $res2= Await-WinRT ($di.Pairing.PairAsync($None))
    if ($res2.Status -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired) { return @{ Status="PairedBasic"; Name=$di.Name } }
    $res3= Await-WinRT ($di.Pairing.PairAsync($Default))
    if ($res3.Status -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired) { return @{ Status="PairedDefault"; Name=$di.Name } }
    return @{ Status="Failed"; Name=$di.Name; Details="$($res.Status)/$($res2.Status)/$($res3.Status)" }
  } finally {
    if($handler){ Unregister-Event -SourceIdentifier $handler.Name | Out-Null }
  }
}
function Pair-ByName([string]$Name,[int]$Retries=3,[int]$Between=2,[string]$Pin){
  $rx = New-Object System.Text.RegularExpressions.Regex ([regex]::Escape($Name)), 'IgnoreCase'
  for($i=1;$i -le $Retries;$i++){
    $d = Find-Devices All | Where-Object { $_.Name -and $rx.IsMatch($_.Name) -and (-not $_.IsPaired) } | Select-Object -First 1
    if($d){ return (Pair-ById -Id $d.Id -Pin $Pin) }
    Start-Sleep -Seconds $Between
  }
  return @{ Status="NoMatch"; Name=$Name }
}
function Format-Mac([UInt64]$a){ ('{0:X12}' -f $a) -replace '(.{2})(?=.)','$1:' }
function Sniffer([int]$Seconds=15,[string]$NameLike){
  Ensure-WinRT
  $w=[Windows.Devices.Bluetooth.Advertisement.BluetoothLEAdvertisementWatcher]::new()
  $sType=[type]::GetType('Windows.Devices.Bluetooth.Advertisement.BluetoothLEScanningMode, Windows, ContentType=WindowsRuntime')
  if($sType){ $Active=[enum]::Parse($sType,'Active'); try{$w.ScanningMode=$Active}catch{} }
  if($NameLike){ $script:Filter = New-Object System.Text.RegularExpressions.Regex ($NameLike),'IgnoreCase' } else { $script:Filter=$null }
  Register-ObjectEvent -InputObject $w -EventName Received -SourceIdentifier 'ADV' -Action {
    $e=$EventArgs; $name=$e.Advertisement.LocalName
    if($script:Filter -and -not $script:Filter.IsMatch([string]$name)){ return }
    $mac=Format-Mac $e.BluetoothAddress
    Write-Host ("ADV  MAC={0,-17} RSSI={1,4}dBm  Name={2}" -f $mac,$e.RawSignalStrengthInDBm, ($name -as [string]))
  } | Out-Null
  try{
    $w.Start()
    Start-Sleep -Seconds $Seconds
    $w.Stop()
  } finally {
    Unregister-Event -SourceIdentifier 'ADV' -ErrorAction SilentlyContinue | Out-Null
  }
}
function Adapters {
  Ensure-WinRT
  $radios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
  $bt = $radios | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }
  $bt | ForEach-Object { [pscustomobject]@{ Name=$_.Name; State=$_.State; Id=$_.DeviceId } } | Sort-Object Name
}
function SetRadio([string]$NameOrId,[string]$State){
  Ensure-WinRT
  $radios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
  $bt = $radios | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }
  $m = $bt | Where-Object { $_.DeviceId -eq $NameOrId -or ($_.Name -and $_.Name -like "*$NameOrId*") } | Select-Object -First 1
  if(-not $m){ throw "No adapter matched '$NameOrId'." }
  $target = if($State -eq 'On'){ [Windows.Devices.Radios.RadioState]::On } else { [Windows.Devices.Radios.RadioState]::Off }
  [void](Await-WinRT ($m.SetStateAsync($target)))
  return @{ Name=$m.Name; State=$m.State }
}
function AdapterInfo {
  Ensure-WinRT
  $a = $null
  try { $a = Await-WinRT ([Windows.Devices.Bluetooth.BluetoothAdapter]::GetDefaultAsync()) } catch {}
  $r = $null
  if($a){ try { $r = Await-WinRT ($a.GetRadioAsync()) } catch {} }
  if(-not $r){
    $rs = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
    $r = $rs | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth } | Select-Object -First 1
  }
  $addr = $null; if($a -and $a.BluetoothAddress){ $addr = ('{0:X12}' -f $a.BluetoothAddress) }
  [pscustomobject]@{
    HasDefaultAdapter = [bool]$a
    BluetoothAddressHex = $addr
    IsLowEnergySupported = $a.IsLowEnergySupported
    IsClassicSupported = $a.IsClassicSupported
    RadioName = $r.Name
    RadioState = $r.State
    RadioId = $r.DeviceId
  }
}

switch($Operation){
  'Scan'        { Find-Devices ($Args.Mode) | ConvertTo-Json -Depth 4; break }
  'PairId'      { (Pair-ById -Id $Args.Id -Pin $Args.Pin) | ConvertTo-Json -Depth 4; break }
  'PairName'    { (Pair-ByName -Name $Args.Name -Retries $Args.Retries -Between $Args.Between -Pin $Args.Pin) | ConvertTo-Json -Depth 4; break }
  'Sniffer'     { Sniffer -Seconds $Args.Seconds -NameLike $Args.NameLike; break }
  'Adapters'    { Adapters | ConvertTo-Json -Depth 4; break }
  'SetRadio'    { (SetRadio -NameOrId $Args.NameOrId -State $Args.State) | ConvertTo-Json -Depth 4; break }
  'AdapterInfo' { AdapterInfo | ConvertTo-Json -Depth 4; break }
  default       { throw "Unknown operation: $Operation" }
}
'@

  # Encode helper + call
  $payload = @{
    Operation = $Operation
    Args      = $Args
  } | ConvertTo-Json -Depth 5

  $runner = @"
`$ErrorActionPreference='Stop'
`$inputJson = @'
$payload
'@ | ConvertFrom-Json
& {
$helper
} -Operation `$inputJson.Operation -Args `$inputJson.Args
"@

  if ($StreamToConsole) {
    # Stream directly (for sniffer/listeners)
    & $ps51 -NoProfile -ExecutionPolicy Bypass -Command $runner
    return
  } else {
    $out = & $ps51 -NoProfile -ExecutionPolicy Bypass -Command $runner 2>&1
    return $out
  }
}

# -------------------- PnP (works in 7.x) --------------------
function Ensure-PnpDeviceModule {
  try { Import-Module PnpDevice -ErrorAction Stop; $script:HasPnpDevice = $true }
  catch { $script:HasPnpDevice = $false }
}
Ensure-PnpDeviceModule

# -------------------- Public cmdlets (PS7 wrappers) --------------------
function Get-BtAdapterInfo {
  $raw = Invoke-PS51 -Operation AdapterInfo
  $obj = $raw | ConvertFrom-Json -Depth 6
  $obj
}
function Get-BtAdapters {
  (Invoke-PS51 -Operation Adapters | ConvertFrom-Json) | Sort-Object Name
}
function Set-BtAdapterState {
  param(
    [Parameter(Mandatory)][ValidateSet('On','Off')]$State,
    [Parameter(Mandatory)][string]$NameOrId
  )
  Invoke-PS51 -Operation SetRadio -Args @{ NameOrId=$NameOrId; State=$State } | Out-Null
  Get-BtAdapters | Where-Object { $_.Name -like "*$NameOrId*" -or $_.Id -eq $NameOrId }
}
function Find-BtDevices {
  param([ValidateSet('BLE','Classic','All')]$Mode='All')
  (Invoke-PS51 -Operation Scan -Args @{ Mode=$Mode } | ConvertFrom-Json) | Sort-Object Name
}
function Pair-BtDevice {
  param([Parameter(Mandatory)][string]$DeviceIdOrAddress,[string]$Pin)
  if ($DeviceIdOrAddress -notlike 'Bluetooth#*') { throw "Provide DeviceInformation.Id starting with 'Bluetooth#'." }
  $resJson = Invoke-PS51 -Operation PairId -Args @{ Id=$DeviceIdOrAddress; Pin=$Pin }
  $res = $resJson | ConvertFrom-Json
  $res
}
function Pair-BtByName {
  param([Parameter(Mandatory)][string]$NameMatch,[string]$Pin,[int]$Retries=6,[int]$BetweenSeconds=2)
  $resJson = Invoke-PS51 -Operation PairName -Args @{ Name=$NameMatch; Retries=$Retries; Between=$BetweenSeconds; Pin=$Pin }
  $res = $resJson | ConvertFrom-Json
  $res
}
function Start-BleSniffer {
  param([int]$Seconds=15,[string]$NameLike)
  Write-Host "Starting live sniffer in a Windows PowerShell 5.1 child..." -ForegroundColor Cyan
  Invoke-PS51 -Operation Sniffer -Args @{ Seconds=$Seconds; NameLike=$NameLike } -StreamToConsole
}
function Get-PairedDevices {
  if (-not $script:HasPnpDevice) {
    Write-Warning "PnpDevice module not available; showing discoverable devices instead."
    return (Find-BtDevices -Mode All | Select-Object @{n='Status';e={'(unknown)'}}, @{n='FriendlyName';e={$_.Name}}, @{n='InstanceId';e={$_.Id}})
  }
  Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
    Where-Object { $_.FriendlyName } |
    Sort-Object FriendlyName |
    Select-Object Status, FriendlyName, InstanceId
}
function Disable-Enable-Device {
  param([Parameter(Mandatory)][string]$InstanceId)
  if (-not $script:HasPnpDevice) { Write-Warning "PnpDevice module not available; cannot toggle '$InstanceId'."; return }
  Disable-PnpDevice -InstanceId $InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
  Start-Sleep 2
  Enable-PnpDevice  -InstanceId $InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
  Write-Host "Toggled device: $InstanceId" -ForegroundColor Yellow
}
function Toggle-BtRadio {
  param([string]$Vid='VID_0BDA')
  if (-not $script:HasPnpDevice) { Write-Warning "PnpDevice module not available; cannot toggle radios."; return }
  $d = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
       Where-Object { $_.InstanceId -match $Vid } | Select-Object -First 1
  if (-not $d) { Write-Error "BT radio with $Vid not found."; return }
  Write-Host "Toggling: $($d.FriendlyName) [$($d.InstanceId)]"
  Disable-Enable-Device -InstanceId $d.InstanceId
}
function Remove-BtDevice {
  param([Parameter(Mandatory)][string]$Match)
  # Try PnP remove first
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
  # Fallback to pnputil
  Write-Warning "Falling back to 'pnputil /remove-device'."
  & pnputil.exe /remove-device "$Match" | Write-Host
}

# Convenience: show adapter + quick tips
function Show-BtStatus {
  Get-BtAdapterInfo | Format-List *
  Write-Host "`nQuick tests:" -ForegroundColor Cyan
  Write-Host "  Start-BleSniffer -Seconds 10" -ForegroundColor DarkGray
  Write-Host "  Find-BtDevices -Mode All | ft" -ForegroundColor DarkGray
}

# -------------------- Hints on load --------------------
Write-Host "Loaded bluetooth-janx-ps7.ps1 (PowerShell 7.5 bridge). Commands:" -ForegroundColor Cyan
@"
Get-BtAdapterInfo
Get-BtAdapters
Set-BtAdapterState -State On -NameOrId '5.3'
Start-BleSniffer -Seconds 15 -NameLike 'Logi'
Find-BtDevices -Mode All | ft
Pair-BtByName -NameMatch 'Keyboard K380' -Retries 8 -BetweenSeconds 1 -Pin '777036'
Pair-BtDevice -DeviceIdOrAddress 'Bluetooth#Bluetooth...'
Get-PairedDevices
Toggle-BtRadio
Remove-BtDevice -Match 'Logitech'
"@ | Write-Host
