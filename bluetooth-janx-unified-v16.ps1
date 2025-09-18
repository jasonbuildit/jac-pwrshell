<# bluetooth-janx-unified-v16.ps1
   - PowerShell 7.5 friendly wrappers + PS 5.1 WinRT worker bridge
   - Robust framed-JSON parsing (safe against chatter)
   - Device scan (BLE/Classic), pair by Id or name, basic BLE sniffer
   - Adapter info, list, toggle
   - Listeners: keyboard by name, any pairable device
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Helper: resolve Windows PowerShell 5.1 path (define only if missing) ---
if (-not (Get-Command Get-Ps51Path -ErrorAction SilentlyContinue)) {
  function Get-Ps51Path {
    $sys32 = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $syswow= Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path $sys32) { return $sys32 }
    if (Test-Path $syswow){ return $syswow }
    throw "Windows PowerShell 5.1 not found."
  }
}

# --- v16 bridge: PS7 -> PS5.1 WinRT worker with framed JSON & safe hashtable marshalling ---
function Invoke-PS51 {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$Operation,  # Scan, PairId, PairName, Adapters, SetRadio, AdapterInfo, Sniffer
    [hashtable]$OpArgs = @{},
    [switch]$StreamToConsole
  )

  $ps51 = Get-Ps51Path

  # ===== Embedded PS5.1 worker (WinRT) =====
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
    # fallback to snapshot if WinRT events cannot be registered
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

# Frame JSON so the PS7 caller can safely parse only the JSON part
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
  Write-Output '<<<JSON_START>>>'
  Write-Output $__json
  Write-Output '<<<JSON_END>>>'
}
'@

  # ===== Build payload in PS7 =====
  $payload = @{ Operation = $Operation; OpArgs = $OpArgs } | ConvertTo-Json -Depth 6

  # Turn PSCustomObject OpArgs into Hashtable for PS5.1 parameter binding
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

  # ===== Run PS5.1 child =====
  $raw = & $ps51 -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $runner 2>&1 | Out-String

  if ($StreamToConsole) {
    $raw | Write-Host
    return
  }

  # Extract framed JSON (if present); else return raw text (e.g., sniffer stream)
  if ($raw -match '(?s)<<<JSON_START>>>(?<json>.*)<<<JSON_END>>>') {
    $json = $Matches['json'].Trim()
    try { return ($json | ConvertFrom-Json -Depth 10) } catch { return $json }
  } else {
    return $raw
  }
}

# ===================== PS7 WRAPPERS (user-friendly) =====================

function Find-BtDevice {
  [CmdletBinding()]
  param([ValidateSet('BLE','Classic','All')][string]$Mode='All')
  Invoke-PS51 -Operation 'Scan' -OpArgs @{ Mode = $Mode }
}

function Pair-BtDevice {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$DeviceIdOrAddress,
    [string]$Pin
  )
  if ($DeviceIdOrAddress -notlike 'Bluetooth#*') {
    throw "In this build, provide DeviceInformation.Id (starts with 'Bluetooth#'). Use Find-BtDevice or the listeners to obtain it."
  }
  Invoke-PS51 -Operation 'PairId' -OpArgs @{ Id = $DeviceIdOrAddress; Pin = $Pin }
}

function Pair-BtByName {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$NameMatch,
    [int]$Retries = 8,
    [int]$BetweenSeconds = 2,
    [string]$Pin
  )
  Invoke-PS51 -Operation 'PairName' -OpArgs @{ Name=$NameMatch; Retries=$Retries; Between=$BetweenSeconds; Pin=$Pin }
}

function Get-BtAdapters {
  Invoke-PS51 -Operation 'Adapters'
}

function Set-BtAdapterState {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][ValidateSet('On','Off')][string]$State,
    [Parameter(Mandatory)][string]$NameOrId
  )
  Invoke-PS51 -Operation 'SetRadio' -OpArgs @{ NameOrId = $NameOrId; State = $State }
}

function Get-BtAdapterInfo {
  Invoke-PS51 -Operation 'AdapterInfo'
}

function Start-BleSniffer {
  [CmdletBinding()]
  param(
    [int]$Seconds = 15,
    [string]$NameLike,
    [ValidateSet('BLE','Classic','All')][string]$Mode='All',
    [switch]$Stream
  )
  if ($Stream) {
    Invoke-PS51 -Operation 'Sniffer' -OpArgs @{ Seconds=$Seconds; NameLike=$NameLike; Mode=$Mode } -StreamToConsole
  } else {
    $text = Invoke-PS51 -Operation 'Sniffer' -OpArgs @{ Seconds=$Seconds; NameLike=$NameLike; Mode=$Mode }
    $text
  }
}

function Listen-ForKeyboard {
  [CmdletBinding()]
  param(
    [string]$NameMatch = 'Keyboard K380',
    [int]$TimeoutMinutes = 5,
    [string]$Pin
  )
  Write-Host ("Listening for '{0}' and attempting auto-pair… (Timeout={1} min)" -f $NameMatch,$TimeoutMinutes) -ForegroundColor Cyan
  $tEnd = (Get-Date).AddMinutes($TimeoutMinutes)
  while ((Get-Date) -lt $tEnd) {
    $scan = Find-BtDevice -Mode All
    $match = $scan | Where-Object { $_.Name -and ($_.Name -like "*$NameMatch*") -and (-not $_.IsPaired) } | Select-Object -First 1
    if ($match) {
      Write-Host ("Found: {0} — pairing..." -f $match.Name) -ForegroundColor Cyan
      $res = Pair-BtDevice -DeviceIdOrAddress $match.Id -Pin $Pin
      return $res
    }
    Start-Sleep -Seconds 1
  }
  Write-Warning ("Finished listening window without a successful pair for '{0}'." -f $NameMatch)
}

function Listen-ForAnyDevice {
  [CmdletBinding()]
  param(
    [int]$Seconds = 15,
    [ValidateSet('BLE','Classic','All')][string]$Mode='All'
  )
  Write-Host ("Listening for ANY pairable Bluetooth device (snapshot loop, {0}s)" -f $Seconds) -ForegroundColor Cyan
  $stop = (Get-Date).AddSeconds($Seconds)
  $printed = New-Object 'System.Collections.Generic.HashSet[string]'
  while ((Get-Date) -lt $stop) {
    $scan = Find-BtDevice -Mode $Mode
    foreach ($d in $scan) {
      if ($d.CanPair -and -not $d.IsPaired -and -not $printed.Contains($d.Id)) {
        [void]$printed.Add($d.Id)
        Write-Host ("PAIRABLE: {0}`n  Id={1}" -f ($d.Name -as [string]), $d.Id) -ForegroundColor Gray
      }
    }
    Start-Sleep -Milliseconds 800
  }
}

# ===================== Load banner =====================
Write-Host "Loaded bluetooth-janx-unified-v16.ps1 (PS7 bridge + 5.1 WinRT). Commands:" -ForegroundColor Green
@(
  'Get-BtAdapterInfo',
  'Get-BtAdapters',
  "Set-BtAdapterState -State On -NameOrId '5.3'",
  'Find-BtDevice -Mode All | ft',
  "Pair-BtByName -NameMatch 'Keyboard K380' -Retries 10 -BetweenSeconds 1 -Pin '777036'",
  "Pair-BtDevice -DeviceIdOrAddress 'Bluetooth#Bluetooth...'",
  "Start-BleSniffer -Seconds 10 -NameLike 'Logi' -Stream",
  "Listen-ForKeyboard -NameMatch 'Keyboard K380' -TimeoutMinutes 3 -Pin '777036'",
  'Listen-ForAnyDevice -Seconds 15',
  'Get-PairedDevices (from your earlier module, if present)'
) | ForEach-Object { Write-Host $_ }
