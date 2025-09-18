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
        @{ Expression = 'Score'; Descending = $true  },
        @{ Expression = 'Name' ; Descending = $false }
      )
}
