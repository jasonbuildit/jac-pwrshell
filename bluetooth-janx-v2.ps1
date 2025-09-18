#Requires -Version 5.1
<# bt51.ps1 — scan (BLE/Classic), pair by MAC (with optional PIN), connect BLE (open GATT),
   list paired, toggle device, and remove paired device. Run as Administrator. #>

param(
  [string]$PreferredAdapterMatch = "5.3",           # prefer your BT 5.3 dongle by name match
  [ValidateSet("BLE","Classic","All")]
  [string]$ScanMode = "All"
)

function Await-WinRT { param($op) $op.AsTask().GetAwaiter().GetResult() }

# --- Adapters ---
function Get-BtAdapters {
  $radios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
  $bt = $radios | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }
  $bt | ForEach-Object {
    [pscustomobject]@{
      Name  = $_.Name
      State = $_.State
      Id    = $_.DeviceId
      Score = if ($_.Name -match [regex]::Escape($PreferredAdapterMatch)) { 100 } else { 0 }
      _raw  = $_
    }
  } | Sort-Object -Property @(
        @{ Expression='Score'; Descending=$true },
        @{ Expression='Name' ; Descending=$false }
      )
}

function Set-BtAdapterState {
  param([ValidateSet("On","Off")]$State,[Parameter(Mandatory)]$Radio)
  if ($State -eq "On") { Await-WinRT ($Radio.SetStateAsync([Windows.Devices.Radios.RadioState]::On)) }
  else { Await-WinRT ($Radio.SetStateAsync([Windows.Devices.Radios.RadioState]::Off)) }
}

# --- Scan ---
function Get-BtAqsFilter {
  param([ValidateSet("BLE","Classic","All")]$Mode = "All")
  switch ($Mode) {
    "BLE"     { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}"' }
    "Classic" { 'System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
    default   { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}" OR System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
  }
}

function Find-BtDevices {
  param(
    [ValidateSet("BLE","Classic","All")]$Mode = "All",
    [int]$Seconds = 10
  )
  $aqs   = Get-BtAqsFilter -Mode $Mode
  $props = @(
    "System.Devices.Aep.DeviceAddress",
    "System.Devices.Aep.IsPaired",
    "System.Devices.Aep.IsConnected",
    "System.Devices.Aep.Bluetooth.Le.IsConnectable"
  )

  # Correct factory:
  $watcher = [Windows.Devices.Enumeration.DeviceInformation]::CreateWatcher(
                $aqs, $props,
                [Windows.Devices.Enumeration.DeviceInformationKind]::AssociationEndpoint
             )

  $results = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new()

  $onAdded = Register-ObjectEvent -InputObject $watcher -EventName Added -Action {
    $di = $EventArgs
    $obj = [pscustomobject]@{
      Name         = $di.Name
      Id           = $di.Id
      Address      = $di.Properties["System.Devices.Aep.DeviceAddress"]
      IsPaired     = $di.Properties["System.Devices.Aep.IsPaired"]
      IsConnected  = $di.Properties["System.Devices.Aep.IsConnected"]
      LeConnectable= $di.Properties["System.Devices.Aep.Bluetooth.Le.IsConnectable"]
    }
    $global:results[$di.Id] = $obj
  }

  $watcher.Start()
  Start-Sleep -Seconds $Seconds
  $watcher.Stop()

  if ($onAdded) { Unregister-Event -SourceIdentifier $onAdded.Name | Out-Null }

  return $results.Values | Sort-Object Name
}

# --- Pair / Connect / PnP helpers ---
function Convert-BtAddrToUlong { param([Parameter(Mandatory)][string]$Address)
  $hex = $Address -replace "[:\-]",""
  if ($hex.Length -ne 12) { throw "Address must be AA:BB:CC:DD:EE:FF" }
  [uint64]::Parse($hex,[System.Globalization.NumberStyles]::HexNumber)
}

function Pair-BtDevice {
  param([Parameter(Mandatory)][string]$DeviceIdOrAddress,[string]$Pin)
  if ($DeviceIdOrAddress -notlike "Bluetooth#*") {
    $m = Find-BtDevices -Mode All -Seconds 8 | Where-Object {
      $_.Address -and ($_.Address -replace "[:\-]" -eq ($DeviceIdOrAddress -replace "[:\-]",""))
    } | Select-Object -First 1
    if (-not $m) { throw "Device not found: $DeviceIdOrAddress" }
    $DeviceIdOrAddress = $m.Id
  }
  $di = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::CreateFromIdAsync($DeviceIdOrAddress))
  if ($di.Pairing.IsPaired) { Write-Host "Already paired: $($di.Name)" -ForegroundColor Yellow; return }
  $custom = $di.Pairing.Custom
  $handler = Register-ObjectEvent -InputObject $custom -EventName PairingRequested -Action {
    $req = $EventArgs
    switch ($req.PairingKind) {
      "ProvidePin" { $p = $using:Pin; if (-not $p) { $p = "0000" }; $req.Accept($p) }
      "ConfirmOnly" { $req.Accept() }
      "DisplayPin" { Write-Host "Enter this PIN on the device: $($req.Pin)" -ForegroundColor Cyan; $req.Accept() }
      "ConfirmPinMatch" { Write-Host "Confirm PIN: $($req.Pin)" -ForegroundColor Cyan; $req.Accept() }
      default { $req.Accept() }
    }
  }
  try {
    $res = Await-WinRT ($custom.PairAsync([Windows.Devices.Enumeration.DevicePairingProtectionLevel]::Default))
    if ($res.Status -ne [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired) { throw "Pair failed: $($res.Status)" }
    Write-Host "Paired: $($di.Name)" -ForegroundColor Green
  } finally { if ($handler) { Unregister-Event -SourceIdentifier $handler.Name | Out-Null } }
}

function Connect-BleDevice {
  param([Parameter(Mandatory)][string]$Address)
  $addrU = Convert-BtAddrToUlong -Address $Address
  $ble = Await-WinRT ([Windows.Devices.Bluetooth.BluetoothLEDevice]::FromBluetoothAddressAsync($addrU))
  if (-not $ble) { throw "BLE device not found/visible: $Address" }
  $svc = Await-WinRT ($ble.GetGattServicesAsync()) # forces link
  Write-Host ("Connected to {0} (Services={1})" -f $ble.Name,$svc.Services.Count) -ForegroundColor Green
  $ble
}

function Get-PairedDevices {
  Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
    Where-Object { $_.FriendlyName } |
    Sort-Object FriendlyName |
    Select-Object Status, FriendlyName, InstanceId
}

function Disable-Enable-Device {
  param([Parameter(Mandatory)][string]$InstanceId)
  Disable-PnpDevice -InstanceId $InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
  Start-Sleep 2
  Enable-PnpDevice  -InstanceId $InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
  Write-Host "Toggled device: $InstanceId" -ForegroundColor Yellow
}

function Remove-BtDevice {
  param([Parameter(Mandatory)][string]$Match)
  $dev = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue | Where-Object {
    $_.InstanceId -eq $Match -or ($_.FriendlyName -and $_.FriendlyName -like "*$Match*")
  } | Select-Object -First 1
  if (-not $dev) { throw "No Bluetooth device matched '$Match'." }
  Write-Host "Removing: $($dev.FriendlyName) [$($dev.InstanceId)]" -ForegroundColor Yellow
  Remove-PnpDevice -InstanceId $dev.InstanceId -Confirm:$false
  Write-Host "Removed." -ForegroundColor Green
}

# --- Prefer BT 5.3 dongle (by name) & turn off others (optional) ---
try {
  $adapters = Get-BtAdapters
  if ($adapters) {
    $primary = $adapters | Select-Object -First 1
    if ($primary.State -ne [Windows.Devices.Radios.RadioState]::On) {
      Set-BtAdapterState -State On -Radio $primary._raw
    }
    foreach ($r in ($adapters | Select-Object -Skip 1)) {
      if ($r.State -eq [Windows.Devices.Radios.RadioState]::On) {
        Write-Host "Turning OFF secondary adapter: $($r.Name)" -ForegroundColor DarkGray
        Set-BtAdapterState -State Off -Radio $r._raw
      }
    }
  }
} catch { Write-Warning "Adapter preference skipped: $($_.Exception.Message)" }

Write-Host "`nReady. Common commands:" -ForegroundColor Cyan
@"
Find-BtDevices -Mode All -Seconds 12 | Format-Table
Pair-BtDevice -DeviceIdOrAddress 'AA:BB:CC:DD:EE:FF'
Pair-BtDevice -DeviceIdOrAddress 'AA:BB:CC:DD:EE:FF' -Pin '0000'
Connect-BleDevice -Address 'AA:BB:CC:DD:EE:FF'
Get-PairedDevices
Disable-Enable-Device -InstanceId 'BTHENUM\DEV_XXXXXXXXXXXX\8&...'
Remove-BtDevice -Match 'WF-1000XM4'   # or pass InstanceId
"@ | Write-Output

