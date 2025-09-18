#Requires -Version 5.1
<#  bluetooth-janx-all.ps1  (Windows PowerShell 5.1)
    - Snapshot scan (BLE/Classic)
    - Pair by MAC or DeviceId (optional PIN)
    - Connect BLE (force link via GATT)
    - List paired, remove (unpair), toggle device
    - Prefer BT 5.3 dongle by name match
    - Listener (polling): waits for a keyboard (e.g., Logitech) to appear, auto-pairs
    Run as Administrator.
#>

param(
  [string]$PreferredAdapterMatch = "5.3",           # prefer dongle whose *name* contains this
  [ValidateSet("BLE","Classic","All")]
  [string]$ScanMode = "All"
)

# ----------------- Helpers -----------------
function Await-WinRT { param($op) $op.AsTask().GetAwaiter().GetResult() }

function New-StringList {
  param([string[]]$Items)
  $list = New-Object 'System.Collections.Generic.List[string]'
  foreach ($s in $Items) { [void]$list.Add($s) }
  return $list
}

# ----------------- Adapters -----------------
function Get-BtAdapters {
  $radios = Await-WinRT ([Windows.Devices.Radios.Radio]::GetRadiosAsync())
  $bt = $radios | Where-Object { $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }
  $items = foreach ($r in $bt) {
    [pscustomobject]@{
      Name  = $r.Name
      State = $r.State
      Id    = $r.DeviceId
      Score = if ($r.Name -match [regex]::Escape($PreferredAdapterMatch)) { 100 } else { 0 }
      _raw  = $r
    }
  }
  $items | Sort-Object -Property @{Expression='Score';Descending=$true}, @{Expression='Name';Descending=$false}
}

function Set-BtAdapterState {
  param([ValidateSet("On","Off")]$State,[Parameter(Mandatory)]$Radio)
  if ($State -eq "On") { Await-WinRT ($Radio.SetStateAsync([Windows.Devices.Radios.RadioState]::On)) }
  else { Await-WinRT ($Radio.SetStateAsync([Windows.Devices.Radios.RadioState]::Off)) }
}

# ----------------- Scanning (snapshot) -----------------
function Get-BtAqsFilter {
  param([ValidateSet("BLE","Classic","All")]$Mode = "All")
  switch ($Mode) {
    "BLE"     { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}"' }
    "Classic" { 'System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
    default   { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}" OR System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
  }
}

function Find-BtDevices {
  param([ValidateSet("BLE","Classic","All")]$Mode="All",[int]$Seconds=8)
  $aqs = Get-BtAqsFilter -Mode $Mode

  # *** 1-arg overload (AQS only) ***
  $coll = ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs)
          ).AsTask().GetAwaiter().GetResult()

  $items = foreach ($di in $coll) {
    [pscustomobject]@{
      Name          = $di.Name
      Id            = $di.Id
      # Properties may be null with 1-arg overload; Pairing.IsPaired still works:
      IsPaired      = (try { [bool]$di.Pairing.IsPaired } catch { $null })
      Address       = $null
      IsConnected   = $null
      LeConnectable = $null
    }
  }
  $items | Sort-Object Name
}

# ----------------- Pair / Connect / PnP -----------------
function Convert-BtAddrToUlong {
  param([Parameter(Mandatory)][string]$Address) # "AA:BB:CC:DD:EE:FF"
  $hex = $Address -replace "[:\-]",""
  if ($hex.Length -ne 12) { throw "Address must be AA:BB:CC:DD:EE:FF" }
  [uint64]::Parse($hex,[System.Globalization.NumberStyles]::HexNumber)
}

function Pair-BtDevice {
  param([Parameter(Mandatory)][string]$DeviceIdOrAddress,[string]$Pin)

  # Allow MAC or DeviceInformation.Id
  if ($DeviceIdOrAddress -notlike "Bluetooth#*") {
    $m = Find-BtDevices -Mode All | Where-Object {
      $_.Address -and ($_.Address -replace "[:\-]" -eq ($DeviceIdOrAddress -replace "[:\-]",""))
    } | Select-Object -First 1
    if (-not $m) { throw "Device not found: $DeviceIdOrAddress" }
    $DeviceIdOrAddress = $m.Id
  }

  $di = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::CreateFromIdAsync($DeviceIdOrAddress))
  if ($di.Pairing.IsPaired) { Write-Host "Already paired: $($di.Name)" -ForegroundColor Yellow; return }

  # PS5.1-friendly capture (avoid $using:)
  $script:PairingPin = $Pin

  $custom = $di.Pairing.Custom
  $handler = Register-ObjectEvent -InputObject $custom -EventName PairingRequested -Action {
    $req = $EventArgs
    switch ($req.PairingKind) {
      "DisplayPin" {
        Write-Host "========== PAIRING PIN ==========" -ForegroundColor Green
        Write-Host "Type this PIN on the KEYBOARD, then press Enter: $($req.Pin)" -ForegroundColor Green
        Write-Host "=================================" -ForegroundColor Green
        $req.Accept()
      }
      "ConfirmPinMatch" { Write-Host "Confirm PIN: $($req.Pin)" -ForegroundColor Green; $req.Accept() }
      "ConfirmOnly"     { $req.Accept() }
      "ProvidePin"      {
        $p = $script:PairingPin; if (-not $p) { $p = "0000" }
        Write-Host "Providing PIN $p" -ForegroundColor Yellow
        $req.Accept($p)
      }
      default           { $req.Accept() }
    }
  }

  try {
    Write-Host "Pairing with: $($di.Name)" -ForegroundColor Cyan
    $res = Await-WinRT ($custom.PairAsync([Windows.Devices.Enumeration.DevicePairingProtectionLevel]::Default))
    if ($res.Status -ne [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired) { throw "Pair failed: $($res.Status)" }
    Write-Host "Paired: $($di.Name)" -ForegroundColor Green
  } finally {
    if ($handler) { Unregister-Event -SourceIdentifier $handler.Name | Out-Null }
    Remove-Variable -Name PairingPin -Scope Script -ErrorAction SilentlyContinue
  }
}

function Connect-BleDevice {
  param([Parameter(Mandatory)][string]$Address)
  $addrU = Convert-BtAddrToUlong -Address $Address
  $ble = Await-WinRT ([Windows.Devices.Bluetooth.BluetoothLEDevice]::FromBluetoothAddressAsync($addrU))
  if (-not $ble) { throw "BLE device not found/visible: $Address" }
  $svc = Await-WinRT ($ble.GetGattServicesAsync()) # opening GATT forces LE link
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

# Convenience: toggle a Realtek radio by VID match (optional)
function Toggle-BtRadio {
  param([string]$Vid = 'VID_0BDA') # Realtek
  $d = Get-PnpDevice -Class Bluetooth -ErrorAction SilentlyContinue |
       Where-Object { $_.InstanceId -match $Vid } | Select-Object -First 1
  if (-not $d) { Write-Error "BT radio with $Vid not found."; return }
  Write-Host "Toggling: $($d.FriendlyName) [$($d.InstanceId)]"
  Disable-Enable-Device -InstanceId $d.InstanceId
}

# ----------------- Listener (polling; PS5.1-safe) -----------------
function Listen-ForKeyboard {
  param(
    [string]$NameMatch = 'Logitech',
    [ValidateSet('BLE','Classic','All')] [string]$Mode = 'All',
    [int]$TimeoutMinutes = 10,
    [string]$Pin  # optional for ProvidePin fallback
  )

  $aqs   = Get-BtAqsFilter -Mode $Mode
  $regex = New-Object System.Text.RegularExpressions.Regex ([regex]::Escape($NameMatch)), 'IgnoreCase'
  $seen  = New-Object 'System.Collections.Generic.HashSet[string]'

  Write-Host "Listening (polling) for '$NameMatch' keyboard… Mode=$Mode, Timeout=$TimeoutMinutes min" -ForegroundColor White
  Write-Host "Put the keyboard in pairing mode now." -ForegroundColor White

  $stopAt = (Get-Date).AddMinutes($TimeoutMinutes)
  while ((Get-Date) -lt $stopAt) {
    try {
      # *** 1-arg overload (AQS only) ***
      $coll = ([Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync($aqs)
              ).AsTask().GetAwaiter().GetResult()

      foreach ($di in $coll) {
        $name = $di.Name
        if (-not $name) { continue }
        if (-not $regex.IsMatch($name)) { continue }

        $isPaired = $false
        try { $isPaired = [bool]$di.Pairing.IsPaired } catch { $isPaired = $false }

        if (-not $isPaired -and -not $seen.Contains($di.Id)) {
          [void]$seen.Add($di.Id)
          Write-Host "Detected candidate: $name — attempting to pair..." -ForegroundColor Cyan
          try { Pair-BtDevice -DeviceIdOrAddress $di.Id -Pin $Pin }
          catch { Write-Warning ("Pair attempt failed for '{0}': {1}" -f $name, $_.Exception.Message) }
        }
      }
    } catch {
      Write-Warning ("Scan error: {0}" -f $_.Exception.Message)
    }
    Start-Sleep -Seconds 1
  }

  Write-Host "Listener finished." -ForegroundColor Yellow
}

# ----------------- Prefer BT 5.3 dongle (optional) -----------------
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

# ----------------- Ready hints -----------------
Write-Host "`nReady. Common commands:" -ForegroundColor Cyan
@"
# Snapshot scan (BLE + Classic)
Find-BtDevices -Mode All | ft

# Pair by MAC (no PIN or confirm-only)
Pair-BtDevice -DeviceIdOrAddress 'AA:BB:CC:DD:EE:FF'

# Pair with fixed PIN (legacy keyboards/headsets)
Pair-BtDevice -DeviceIdOrAddress 'AA:BB:CC:DD:EE:FF' -Pin '0000'

# Force BLE connect (opens GATT)
Connect-BleDevice -Address 'AA:BB:CC:DD:EE:FF'

# List paired devices
Get-PairedDevices

# Disconnect/reconnect a stubborn device (InstanceId from above)
Disable-Enable-Device -InstanceId 'BTHENUM\DEV_XXXXXXXXXXXX\8&...'

# Remove (unpair) by name or InstanceId
Remove-BtDevice -Match 'Logitech Keyboard'

# Listen for a Logitech keyboard and auto-pair (10 min timeout)
Listen-ForKeyboard -NameMatch 'Logitech' -Mode All -TimeoutMinutes 10

# If your model is BLE-only:
Listen-ForKeyboard -NameMatch 'Logitech' -Mode BLE -TimeoutMinutes 10
"@ | Write-Output
