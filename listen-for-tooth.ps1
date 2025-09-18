#Requires -Version 5.1
<# 
  Listens for a Logitech Bluetooth keyboard entering pairing mode and auto-pairs.
  Run in Windows PowerShell 5.1 as Administrator.
#>

param(
  [string]$NameMatch = 'Logitech',
  [ValidateSet('BLE','Classic','All')] [string]$ScanMode = 'All',
  [int]$TimeoutMinutes = 10,
  [string]$Pin  # optional fallback for ProvidePin, e.g., '0000'
)

# Simple await for WinRT in PS 5.1
function Await-WinRT { param($op) $op.AsTask().GetAwaiter().GetResult() }

# AQS filter for BLE/Classic/All
function Get-BtAqsFilter {
  param([ValidateSet('BLE','Classic','All')]$Mode = 'All')
  switch ($Mode) {
    'BLE'     { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}"' }
    'Classic' { 'System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
    default   { 'System.Devices.Aep.ProtocolId:="{bb7bb05e-5972-42b5-94fc-76eaa7084d49}" OR System.Devices.Aep.ProtocolId:="{e0cbf06c-cd8b-4647-bb8a-263B43F0F974}"' }
  }
}

# Kick off pairing for a discovered device
function Start-Pairing {
  param([Parameter(Mandatory)][string]$DeviceId)

  $di = Await-WinRT ([Windows.Devices.Enumeration.DeviceInformation]::CreateFromIdAsync($DeviceId))
  if (-not $di) { Write-Warning "DeviceInformation null for: $DeviceId"; return }
  if ($di.Pairing.IsPaired) { Write-Host "Already paired: $($di.Name)" -ForegroundColor Yellow; return }

  $custom = $di.Pairing.Custom

  $handler = Register-ObjectEvent -InputObject $custom -EventName PairingRequested -Action {
    $req = $EventArgs
    switch ($req.PairingKind) {
      'DisplayPin' {
        # Host should display the PIN; type it on the keyboard and press Enter.
        Write-Host "========== PAIRING PIN ==========" -ForegroundColor Green
        Write-Host "Type this PIN on the KEYBOARD, then press Enter: $($req.Pin)" -ForegroundColor Green
        Write-Host "=================================" -ForegroundColor Green
        $req.Accept()
      }
      'ConfirmPinMatch' {
        Write-Host "Confirm PIN shown on device: $($req.Pin)" -ForegroundColor Green
        $req.Accept()
      }
      'ConfirmOnly' {
        $req.Accept()
      }
      'ProvidePin' {
        $p = $using:Pin; if (-not $p) { $p = '0000' }
        Write-Host "Providing PIN $p" -ForegroundColor Yellow
        $req.Accept($p)
      }
      default {
        $req.Accept()
      }
    }
  }

  try {
    Write-Host "Pairing with: $($di.Name)" -ForegroundColor Cyan
    $res = Await-WinRT ($custom.PairAsync([Windows.Devices.Enumeration.DevicePairingProtectionLevel]::Default))
    if ($res.Status -eq [Windows.Devices.Enumeration.DevicePairingResultStatus]::Paired) {
      Write-Host "Paired successfully with: $($di.Name)" -ForegroundColor Green
    } else {
      Write-Warning "Pair failed: $($res.Status)"
    }
  } finally {
    if ($handler) { Unregister-Event -SourceIdentifier $handler.Name | Out-Null }
  }
}

# Build watcher
$aqs = Get-BtAqsFilter -Mode $ScanMode
$props = @(
  'System.Devices.Aep.DeviceAddress',
  'System.Devices.Aep.IsPaired',
  'System.Devices.Aep.IsConnected'
)
$watcher = [Windows.Devices.Enumeration.DeviceInformation]::CreateWatcher(
  $aqs, $props, [Windows.Devices.Enumeration.DeviceInformationKind]::AssociationEndpoint
)

$targetRegex = [regex]::new([regex]::Escape($NameMatch), 'IgnoreCase')
$pairingQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Added & Updated handlers
$onAdded = Register-ObjectEvent -InputObject $watcher -EventName Added -Action {
  $di = $EventArgs
  try {
    $name = $di.Name
    $paired = $di.Properties['System.Devices.Aep.IsPaired']
    if ($name -and $using:targetRegex.IsMatch($name) -and -not $paired) {
      Write-Host "Detected candidate: $name" -ForegroundColor Cyan
      $using:pairingQueue.Enqueue($di.Id) | Out-Null
    }
  } catch {}
}

$onUpdated = Register-ObjectEvent -InputObject $watcher -EventName Updated -Action {
  $ud = $EventArgs
  try {
    $name = $ud.Properties?['System.ItemNameDisplay']
    $paired = $ud.Properties?['System.Devices.Aep.IsPaired']
    if ($name -and $using:targetRegex.IsMatch($name) -and ($paired -eq $false)) {
      Write-Host "Updated candidate: $name" -ForegroundColor Cyan
      $using:pairingQueue.Enqueue($ud.Id) | Out-Null
    }
  } catch {}
}

# Start listening
Write-Host "Listening for '$NameMatch' keyboard… (Mode: $ScanMode, Timeout: $TimeoutMinutes min)" -ForegroundColor White
Write-Host "Put the keyboard in pairing mode now." -ForegroundColor White
$watcher.Start()

$stopAt = (Get-Date).AddMinutes($TimeoutMinutes)
try {
  while ((Get-Date) -lt $stopAt) {
    # Dequeue and try to pair
    if ($pairingQueue.TryDequeue([ref]$nextId)) {
      try { Start-Pairing -DeviceId $nextId } catch { Write-Warning $_.Exception.Message }
    } else {
      Start-Sleep -Milliseconds 250
    }
  }
} finally {
  $watcher.Stop()
  foreach ($h in @($onAdded,$onUpdated)) {
    if ($h) { Unregister-Event -SourceIdentifier $h.Name | Out-Null }
  }
}

Write-Host "Listener finished." -ForegroundColor Yellow
