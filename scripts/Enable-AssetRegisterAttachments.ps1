$ErrorActionPreference = 'Stop'
$lists = @(
  'AssetRegister_Hardware_Laptop',
  'AssetRegister_NetworkDevices',
  'AssetRegister_AssetTracker',
  'AssetRegister_OutsourcedServices',
  'AssetRegister_Software',
  'AssetRegister_Hardware_Server',
  'AssetRegister_Cloud_Providers',
  'AssetRegister_Hardware_Accessories',
  'AssetRegister_Intellectual_Property',
  'AssetRegister_MediaDevices',
  'AssetRegister_Mobile_Devices',
  'AssetRegister_Client_Accessories'
)
foreach ($list in $lists) {
  Set-PnPList -Identity $list -EnableAttachments $true
  Write-Host "Enabled attachments: $list"
}
