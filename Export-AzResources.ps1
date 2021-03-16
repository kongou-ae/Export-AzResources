$ErrorActionPreference = "stop"
Add-Type -Assembly System.Drawing
$bgc =  [System.Drawing.Color]::FromArgb(217,225,242)

#-------------------------------------------------------------------------------------------
# Gather information which will be reported by this script 
#-------------------------------------------------------------------------------------------
$vms = Get-AzVM
$vnets = Get-AzVirtualNetwork
$disks = Get-AzDisk
$nics = Get-AzNetworkInterface
$nsgs = Get-AzNetworkSecurityGroup
$pips = Get-AzPublicIpAddress

#-------------------------------------------------------------------------------------------
# Initialize
#-------------------------------------------------------------------------------------------
# Load excel files.


$templateFileName = "$PSScriptRoot/azReportTemplate.xlsx"
$fileName = "$PSScriptRoot/azReport.xlsx"

if( Test-Path $fileName ){
    Remove-Item $fileName
}

# Creat default sheets
Export-Excel $fileName -WorksheetName "README" 
Export-Excel $fileName -WorksheetName "SUMMARY"

# Load default sheets
$excelPackage = Open-ExcelPackage -Path $fileName
$templatePackage = Open-ExcelPackage -Path $templateFileName

#-------------------------------------------------------------------------------------------
# Create README
#-------------------------------------------------------------------------------------------
$readmeMsg = @(
    "",
    "Thanks for using Export-AzResources.",
    "If you find any issue or any request, could you open issue to https://github.com/kongou-ae/Export-AzResources, please?",
    "",
    "@kongou_ae"
)

$readmeWs = $excelPackage.Workbook.Worksheets["README"]
for($i=1;$i -le $readmeMsg.Count; $i++){
    Set-ExcelRange -Worksheet $readmeWs -Range "A${i}" -Value $readmeMsg[$i] -FontSize 16
}

#-------------------------------------------------------------------------------------------
# Create SUMMARY
#-------------------------------------------------------------------------------------------
$summaryWs = $excelPackage.Workbook.Worksheets["SUMMARY"]
Set-ExcelRange -Worksheet $summaryWs -Range "A1" -Value "Export-AzResources" -FontSize 16 -Bold

# create the summary of VirtualMachine
Set-ExcelRange -Worksheet $summaryWs -Range "A3" -Value "VirtualMachines" -FontSize 12 -Bold

$row = 4
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "VmSize" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
$row ++

for($i=0;$i -lt $vms.Count; $i++){
    $vm = $vms[$i]
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $vm.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $vm.ResourceGroupName -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $vm.Location -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value $vm.HardwareProfile.VmSize -BorderAround Thin

    $formular = '=HYPERLINK("#VirtualMachines!vm_' + ($($vm.Name) -replace "-","_") + '","Link")'
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Formula $formular -BorderAround Thin
    $row++
}

$row += 1

# create the summary of VirtualNetwork
Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "VirtualNetwork" -FontSize 12 -Bold
$row ++
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "AddressPrefixes" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
$row ++

for($i=0;$i -lt $vnets.Count; $i++){
    $vnet = $vnets[$i]
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $vnet.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $vnet.ResourceGroupName -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $vnet.Location -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value ($vnet.AddressSpace.AddressPrefixes -join ",") -BorderAround Thin

    $formular = '=HYPERLINK("#VirtualNetworks!vnet_' + ($($vnet.Name) -replace "-","_") + '","Link")'
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Formula $formular -BorderAround Thin
    $row++
}

$row += 1

# create the summary of Disk
Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "Disk" -FontSize 12 -Bold
$row ++
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "Sku" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "Size" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
$row ++

for($i=0;$i -lt $disks.Count; $i++){
    $disk = $disks[$i]
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $disk.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $disk.ResourceGroupName -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $disk.Location -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value $disk.Sku.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value $disk.DiskSizeGB -BorderAround Thin

    $formular = '=HYPERLINK("#disks!vnet_' + ($($disk.Name) -replace "-","_") + '","Link")'
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Formula $formular -BorderAround Thin
    $row++
}

$row += 1

# create the summary of Nic
Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "Network Interface" -FontSize 12 -Bold
$row ++
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "PrivateIpAddress" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "PrivateIpAllocationMethod" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
$row ++

for($i=0;$i -lt $nics.Count; $i++){
    $nic = $nics[$i]
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $nic.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $nic.ResourceGroupName -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $nic.Location -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value $nic.IpConfigurations[0].PrivateIpAddress -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value $nic.IpConfigurations[0].PrivateIpAllocationMethod -BorderAround Thin

    $formular = '=HYPERLINK("#nics!nic_' + ($($nic.Name) -replace "-","_") + '","Link")'
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Formula $formular -BorderAround Thin
    $row++
}

$row += 1

# create the summary of Network Security Group
Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "Network Security Group" -FontSize 12 -Bold
$row ++
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
$row ++

for($i=0;$i -lt $nsgs.Count; $i++){
    $nsg = $nsgs[$i]
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $nsg.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $nsg.ResourceGroupName -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $nsg.Location -BorderAround Thin


    $formular = '=HYPERLINK("#nsgs!nsg_' + ($($nsg.Name) -replace "-","_") + '","Link")'
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Formula $formular -BorderAround Thin
    $row++
}

# create the summary of PublicIP Address
Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "Public Ip Address" -FontSize 12 -Bold
$row ++
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "PublicIpAllocationMethod" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "IpAddress" -Width 20 -BackgroundColor $bgc -BorderAround Thin
Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
$row ++

for($i=0;$i -lt $pips.Count; $i++){
    $pip = $pips[$i]
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $pip.Name -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $pip.ResourceGroupName -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $pip.Location -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value $pip.PublicIpAllocationMethod -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value $pip.IpAddress -BorderAround Thin

    $formular = '=HYPERLINK("#pips!pip_' + ($($pip.Name) -replace "-","_") + '","Link")'
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Formula $formular -BorderAround Thin
    $row++
}

#-------------------------------------------------------------------------------------------
# Create VirtualMachine
#-------------------------------------------------------------------------------------------
if ( $vms -ne $Null ){
    $ws = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "VirtualMachines"
    $shortCols = @("A","B","C","D","E")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $ws -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
    }
    Set-ExcelRange -Worksheet $ws -Range "F:F" -Width (20).ToString()
    Set-ExcelRange -Worksheet $ws -Range "G:G" -Width (100).ToString()
}

$vmWinHeight = 60
$vmLinuxHeight = 60
$workingRow = 1 # エクスポート中のリソースのスタート行
for($i = 0; $i -lt $vms.Count; $i++){
    $vm = $vms[$i]

    Write-Output "Exporting $($vm.Name)"
    if ( $vm.StorageProfile.OsDisk.OsType -eq "Windows" ){
        $templatePackage.Workbook.Worksheets["VirtualMachine_windows"].Cells["A1:G60"].Copy($ws.Cells["A${workingRow}:G$($workingRow + $vmWinHeight)"])          
    }

    if ( $vm.StorageProfile.OsDisk.OsType -eq "Linux" ){
        $templatePackage.Workbook.Worksheets["VirtualMachine_linux"].Cells["A1:G60"].Copy($ws.Cells["A${workingRow}:G$($workingRow + $vmLinuxHeight)"])          
    }

    Set-ExcelRange -Worksheet $ws -Range "A$($workingRow)" -Value $vm.Name -Bold
    Add-ExcelName -Range $ws.Cells["A$($workingRow)"] -RangeName "vm_$($vm.Name)" -WarningAction SilentlyContinue
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 1)" -Value $vm.ResourceGroupName
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 2)" -Value $vm.Name
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 3)" -Value $vm.Location
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 4)" -Value $vm.LicenseType
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 7)" -Value $vm.DiagnosticsProfile.BootDiagnostics.Enabled
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 8)" -Value $vm.DiagnosticsProfile.BootDiagnostics.StorageUri
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 10)" -Value $vm.HardwareProfile.VmSize
    # ToDo: Support multiple nics
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 14)" -Value $vm.NetworkProfile.NetworkInterfaces[0].Primary
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 15)" -Value $vm.NetworkProfile.NetworkInterfaces[0].id
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 17)" -Value $vm.OSProfile.ComputerName
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 18)" -Value $vm.OSProfile.AdminUsername
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 19)" -Value $vm.OSProfile.WindowsConfiguration.ProvisionVMAgent
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 22)" -Value $vm.OSProfile.WindowsConfiguration.EnableAutomaticUpdates
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 23)" -Value $vm.OSProfile.WindowsConfiguration.TimeZone
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 24)" -Value $vm.OSProfile.WindowsConfiguration.AdditionalUnattendContent
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 26)" -Value $vm.OSProfile.WindowsConfiguration.PatchSettings.PatchMode
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 27)" -Value $vm.OSProfile.WindowsConfiguration.PatchSettings.EnableHotpatching
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 28)" -Value $vm.OSProfile.WindowsConfiguration.WinRM
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 29)" -Value $vm.OSProfile.Secrets
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 30)" -Value $vm.OSProfile.AllowExtensionOperations
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 31)" -Value $vm.BillingProfile
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 32)" -Value $vm.Plan
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 35)" -Value $vm.StorageProfile.ImageReference.Publisher
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 36)" -Value $vm.StorageProfile.ImageReference.Offer
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 37)" -Value $vm.StorageProfile.ImageReference.Sku
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 38)" -Value $vm.StorageProfile.ImageReference.Version
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 39)" -Value $vm.StorageProfile.ImageReference.ExactVersion
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 40)" -Value $vm.StorageProfile.ImageReference.id
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 42)" -Value $vm.StorageProfile.OsDisk.OsType
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 43)" -Value $vm.StorageProfile.OsDisk.EncryptionSettings
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 44)" -Value $vm.StorageProfile.OsDisk.Name
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 45)" -Value $vm.StorageProfile.OsDisk.Caching
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 46)" -Value $vm.StorageProfile.OsDisk.WriteAcceleratorEnabled
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 47)" -Value $vm.StorageProfile.OsDisk.CreateOption
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 49)" -Value $vm.StorageProfile.OsDisk.ManagedDisk.StorageAccountType
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 50)" -Value $vm.StorageProfile.OsDisk.ManagedDisk.DiskEncryptionSet
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 51)" -Value $vm.StorageProfile.OsDisk.ManagedDisk.id
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 57)" -Value $vm.EvictionPolicy
    Set-ExcelRange -Worksheet $ws -Range "G$($workingRow + 58)" -Value $vm.Priority

    for($j=0;$j -lt $vm.StorageProfile.DataDisks.Count;$j++ ){
        Write-Output "Adding $($vm.StorageProfile.DataDisks[$j].Name) to $($vm.Name)"
        $addedRowNumbers = 11 # 足される行数
        $addedRowPoint = 54 # リソースの中の何行目に足されるか
        $fromRow = ($workingRow - 1) + $addedRowPoint + $j * $addedRowNumbers # 挿入が始まる行番号
        $toRow = ($workingRow -1) + $addedRowPoint + ($j + 1) * $addedRowNumbers -1 # 挿入が終わる行番号
        $ws.InsertRow($fromRow,$addedRowNumbers) 
        $templatePackage.Workbook.Worksheets["DataDisk"].Cells["A1:G${addedRowNumbers}"].Copy($ws.Cells["A${fromRow}:G${toRow}"])

        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 1)" -Value  $vm.StorageProfile.DataDisks[$j].Lun
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 2)" -Value  $vm.StorageProfile.DataDisks[$j].Name
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 3)" -Value  $vm.StorageProfile.DataDisks[$j].Caching
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 4)" -Value  $vm.StorageProfile.DataDisks[$j].CreateOption
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 6)" -Value  $vm.StorageProfile.DataDisks[$j].ManagedDisk.StorageAccountType
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 7)" -Value  $vm.StorageProfile.DataDisks[$j].ManagedDisk.DiskEncryptionSet
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 8)" -Value  $vm.StorageProfile.DataDisks[$j].ManagedDisk.Id
        Set-ExcelRange -Worksheet $ws -Range "G$($fromRow + 9)" -Value  $vm.StorageProfile.DataDisks[$j].ToBeDetached
    }

    $workingRow += $addedRowNumbers * $vm.StorageProfile.DataDisks.Count

    if ( $vm.StorageProfile.OsDisk.OsType -eq "Windows" ){
        $workingRow += $vmWinHeight + 1
    }
    if ( $vm.StorageProfile.OsDisk.OsType -eq "Linux" ){
        $workingRow += $vmLinuxHeight + 1
    }
}

#-------------------------------------------------------------------------------------------
# Create VirtualNetwork
#-------------------------------------------------------------------------------------------

if ( $vnets -ne $Null ){
    $vnetWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "VirtualNetworks"
    $shortCols = @("A","B","C","D","E")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $vnetWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
    }
    Set-ExcelRange -Worksheet $vnetWs -Range "F:F" -Width (20).ToString()
    Set-ExcelRange -Worksheet $vnetWs -Range "G:G" -Width (100).ToString()
}

$vnetHeight = 12
$workingRow = 1
for($i = 0; $i -lt $vnets.Count; $i++){
    $vnet = $vnets[$i]

    Write-Output "Exporting $($vnet.Name)"

    $templatePackage.Workbook.Worksheets["VirtualNetwork"].Cells["A1:G${vnetHeight}"].Copy($vnetWs.Cells["A${workingRow}:G$($workingRow + $vnetHeight)"])

    Set-ExcelRange -Worksheet $vnetWs -Range "A$($workingRow)" -Value $vnet.Name -Bold
    Add-ExcelName -Range $vnetWs.Cells["A$($workingRow)"] -RangeName "vnet_$($vnet.Name)" -WarningAction SilentlyContinue
    Set-ExcelRange -Worksheet $vnetWs -Range "G$($workingRow + 1)" -Value $vnet.ResourceGroupName
    Set-ExcelRange -Worksheet $vnetWs -Range "G$($workingRow + 2)" -Value $vnet.Name
    Set-ExcelRange -Worksheet $vnetWs -Range "G$($workingRow + 3)" -Value $vnet.Location
    Set-ExcelRange -Worksheet $vnetWs -Range "G$($workingRow + 5)" -Value ($vnet.AddressSpace.AddressPrefixes -join ",")
    Set-ExcelRange -Worksheet $vnetWs -Range "G$($workingRow + 7)" -Value ($vnet.DhcpOptions.DnsServers -join ",")

    for($j=0;$j -lt $vnet.Subnets.Count;$j++ ){
        Write-Output "Adding $($vnet.Subnets[$j].Name) to $($vnet.Name)"
        $addedRowNumbers = 15 # 足される行数
        $addedRowPoint = 10 # リソースの中の何行目に足されるか
        $fromRow = ($workingRow - 1) + $addedRowPoint + $j * $addedRowNumbers # 挿入が始まる行番号
        $toRow = ($workingRow -1) + $addedRowPoint + ($j + 1) * $addedRowNumbers -1 # 挿入が終わる行番号
        $vnetWs.InsertRow($fromRow,$addedRowNumbers) 
        $templatePackage.Workbook.Worksheets["Subnet"].Cells["A1:G${addedRowNumbers}"].Copy($vnetWs.Cells["A${fromRow}:G${toRow}"])

        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 1)" -Value  $vnet.Subnets[$j].Name
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 2)" -Value  ($vnet.Subnets[$j].AddressPrefix -join ",")
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 3)" -Value  $vnet.Subnets[$j].ServiceAssociationLinks
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 5)" -Value  $vnet.Subnets[$j].NetworkSecurityGroup.Id
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 7)" -Value  $vnet.Subnets[$j].RouteTable.DisableBgpRoutePropagation
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 8)" -Value  $vnet.Subnets[$j].RouteTable.Id
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 10)" -Value  $vnet.Subnets[$j].NatGateway.Id
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 11)" -Value  $vnet.Subnets[$j].ServiceEndpoints
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 12)" -Value  $vnet.Subnets[$j].ServiceEndpointPolicies
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 13)" -Value  $vnet.Subnets[$j].PrivateLinkServiceNetworkPolicies
        Set-ExcelRange -Worksheet $vnetWs -Range "G$($fromRow + 14)" -Value  $vnet.Subnets[$j].Delegations
        

    }

    $workingRow += $addedRowNumbers * $vnet.Subnets.Count

    $workingRow += $vnetHeight + 1
}



#-------------------------------------------------------------------------------------------
# Create Disk
#-------------------------------------------------------------------------------------------

if ( $vnets -ne $Null ){
    $diskWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "Disks"
    $shortCols = @("A","B","C","D","E")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $diskWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
    }
    Set-ExcelRange -Worksheet $diskWs -Range "F:F" -Width (20).ToString()
    Set-ExcelRange -Worksheet $diskWs -Range "G:G" -Width (100).ToString()
    
}

$diskHeight = 23
$workingRow = 1
for($i = 0; $i -lt $disks.Count; $i++){
    $disk = $disks[$i]

    Write-Output "Exporting $($disk.Name)"

    $templatePackage.Workbook.Worksheets["Disk"].Cells["A1:G${diskHeight}"].Copy($diskWs.Cells["A${workingRow}:G$($workingRow + $diskHeight)"])

    Set-ExcelRange -Worksheet $diskWs -Range "A$($workingRow)" -Value $disk.Name -Bold
    Add-ExcelName -Range $diskWs.Cells["A$($workingRow)"] -RangeName "disk_$($disk.Name)" -WarningAction SilentlyContinue
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 1)" -Value $disk.ResourceGroupName
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 2)" -Value $disk.Name
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 3)" -Value $disk.Location
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 4)" -Value $disk.ManagedBy
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 6)" -Value $disk.sku.Name
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 7)" -Value $disk.sku.Tier
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 8)" -Value $disk.Zone
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 10)" -Value $disk.CreationData.CreateOption
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 11)" -Value $disk.CreationData.StorageAccountId
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 12)" -Value $disk.CreationData.ImageReference
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 13)" -Value $disk.CreationData.GalleryImageReference
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 14)" -Value $disk.DiskSizeGB
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 15)" -Value $disk.DiskState
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 17)" -Value $disk.Encryption.DiskEncryptionSetId
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 18)" -Value $disk.Encryption.Type
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 19)" -Value $disk.ShareInfo
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 20)" -Value $disk.NetworkAccessPolicy
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 21)" -Value $disk.Tier
    Set-ExcelRange -Worksheet $diskWs -Range "G$($workingRow + 22)" -Value $disk.BurstingEnabled

    $workingRow += $diskHeight + 1
}


#-------------------------------------------------------------------------------------------
# Create Nic
#-------------------------------------------------------------------------------------------

if ( $nics -ne $Null ){
    Write-Output "Adding the new worksheet for Network Interface"
    $nicWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "Nics"
    $shortCols = @("A","B","C","D","E")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $nicWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
    }
    Set-ExcelRange -Worksheet $nicWs -Range "F:F" -Width (20).ToString()
    Set-ExcelRange -Worksheet $nicWs -Range "G:G" -Width (100).ToString()
}

$nicHeight = 28
$workingRow = 1
for($i = 0; $i -lt $nics.Count; $i++){
    $nic = $nics[$i]

    Write-Output "Exporting $($nic.Name)"

    $templatePackage.Workbook.Worksheets["Nic"].Cells["A1:G${nicHeight}"].Copy($nicWs.Cells["A${workingRow}:G$($workingRow + $nicHeight)"])

    Set-ExcelRange -Worksheet $nicWs -Range "A$($workingRow)" -Value $nic.Name -Bold
    Add-ExcelName -Range $nicWs.Cells["A$($workingRow)"] -RangeName "nic_$($nic.Name)" -WarningAction SilentlyContinue
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 1)" -Value $nic.ResourceGroupName
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 2)" -Value $nic.Name
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 3)" -Value $nic.Location
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 4)" -Value $nic.VirtualMachine
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 7)" -Value $nic.IpConfigurations[0].Name
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 8)" -Value $nic.IpConfigurations[0].PrivateIpAddress
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 9)" -Value $nic.IpConfigurations[0].PrivateIpAddressVersion
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 11)" -Value $nic.IpConfigurations[0].Subnet.Id
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 13)" -Value $nic.IpConfigurations[0].PublicIpAddress.Id
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 14)" -Value $nic.IpConfigurations[0].PrivateIpAddressVersion
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 15)" -Value $nic.IpConfigurations[0].LoadBalancerBackendAddressPools
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 16)" -Value $nic.IpConfigurations[0].Primary
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 17)" -Value $nic.IpConfigurations[0].ApplicationGatewayBackendAddressPools
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 18)" -Value $nic.IpConfigurations[0].ApplicationSecurityGroups
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 20)" -Value $nic.DnsSettings.DnsServers
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 21)" -Value $nic.DnsSettings.AppliedDnsServers
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 22)" -Value $nic.EnableIPForwarding
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 23)" -Value $nic.EnableAcceleratedNetworking
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 25)" -Value $nic.NetworkSecurityGroup.Id
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 16)" -Value $nic.Primary
    Set-ExcelRange -Worksheet $nicWs -Range "G$($workingRow + 27)" -Value $nic.MacAddress

    $workingRow += $nicHeight + 1
}


#-------------------------------------------------------------------------------------------
# Create Nsg
#-------------------------------------------------------------------------------------------

if ( $nsgs -ne $Null ){
    Write-Output "Adding the new worksheet for Network interface"
    $nsgWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "Nsgs"
    $shortCols = @("A","B","C","D","E")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $nsgWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
    }
    $shortCols = @("F","G","H","I","J","K","L","M","P")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $nsgWs -Range "${shortCol}:${shortCol}" -Width (20).ToString()
    }
}

$nsgHeight = 9
$workingRow = 1
for($i = 0; $i -lt $nsgs.Count; $i++){
    $nsg = $nsgs[$i]

    Write-Output "Exporting $($nsg.Name)"

    $templatePackage.Workbook.Worksheets["nsg"].Cells["A1:P${nsgHeight}"].Copy($nsgWs.Cells["A${workingRow}:P$($workingRow + $nsgHeight)"])

    Set-ExcelRange -Worksheet $nsgWs -Range "A$($workingRow)" -Value $nsg.Name -Bold
    Add-ExcelName -Range $nsgWs.Cells["A$($workingRow)"] -RangeName "nsg_$($nsg.Name)" -WarningAction SilentlyContinue
    Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow + 1)" -Value $nsg.ResourceGroupName
    Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow + 2)" -Value $nsg.Name
    Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow + 3)" -Value $nsg.Location
    $workingRow += 6

    $nsgWs.InsertRow($workingRow,$nsg.SecurityRules.Count)
    $sortedNsgRules = $nsg.SecurityRules | Sort-Object Direction,Priority
    for($j=0;$j -lt $nsg.SecurityRules.Count; $j++){
        Set-ExcelRange -Worksheet $nsgWs -Range "A${workingRow}" -Value $sortedNsgRules[$j].Direction -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "A${workingRow}:E${workingRow}" -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "F${workingRow}" -Value $sortedNsgRules[$j].Priority -BorderAround Thick -BackgroundColor white   
        Set-ExcelRange -Worksheet $nsgWs -Range "G${workingRow}" -Value $sortedNsgRules[$j].Name -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "H${workingRow}" -Value $sortedNsgRules[$j].Protocol -BorderAround Thick -BackgroundColor white   
        Set-ExcelRange -Worksheet $nsgWs -Range "I${workingRow}" -Value $sortedNsgRules[$j].SourceAddressPrefix -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "J${workingRow}" -Value $sortedNsgRules[$j].SourceApplicationSecurityGroups -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "K${workingRow}" -Value $sortedNsgRules[$j].SourcePortRange -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "L${workingRow}" -Value $sortedNsgRules[$j].DestinationAddressPrefix -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "M${workingRow}" -Value $sortedNsgRules[$j].DestinationApplicationSecurityGroups -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "N${workingRow}" -Value $sortedNsgRules[$j].DestinationPortRange -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "O${workingRow}" -Value $sortedNsgRules[$j].Access -BorderAround Thick -BackgroundColor white
        Set-ExcelRange -Worksheet $nsgWs -Range "P${workingRow}" -Value $sortedNsgRules[$j].Description -BorderAround Thick -BackgroundColor white
        $workingRow ++ 
    }

    $nsg.NetworkInterfaces | ForEach-Object {
        $workingRow ++ # 空の行の位置に移動
        $k = 0
        $nsgWs.InsertRow($workingRow,2) ## 空行とID行を追加
        $templatePackage.Workbook.Worksheets["Nsg1stArrayOnlyId"].Cells["A1:K2"].Copy($nsgWs.Cells["A${workingRow}:P$($workingRow + 1)"])
        $workingRow ++ # ID 行に移動
        Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow)" -Value $nsg.NetworkInterfaces[$k].Id -BorderAround Thick -BackgroundColor white
        $k++
    }

    $workingRow ++ 
    $nsg.Subnets | ForEach-Object {
        $workingRow ++ # 空の行の位置に移動
        $k = 0
        $nsgWs.InsertRow($workingRow,2)
        $templatePackage.Workbook.Worksheets["Nsg1stArrayOnlyId"].Cells["A1:K2"].Copy($nsgWs.Cells["A${workingRow}:P$($workingRow + 1)"])
        $workingRow ++ 
        Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow)" -Value $nsg.Subnets[$k].Id -BorderAround Thick -BackgroundColor white
        $k++

    }
    Set-ExcelRange -Worksheet $nsgWs -Range "A$($workingRow):G$($workingRow)" -BorderBottom Thick

    $workingRow += 2 # NetworkInterfaces,Subnets, space の3つ
}


#-------------------------------------------------------------------------------------------
# Create pip
#-------------------------------------------------------------------------------------------

if ( $pips -ne $Null ){
    Write-Output "Adding the new worksheet for Network Interface"
    $pipWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "pips"
    $shortCols = @("A","B","C","D","E")
    foreach ($shortCol in $shortCols) {
        Set-ExcelRange -Worksheet $pipWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
    }
    Set-ExcelRange -Worksheet $pipWs -Range "F:F" -Width (20).ToString()
    Set-ExcelRange -Worksheet $pipWs -Range "G:G" -Width (100).ToString()
}

$pipHeight = 16
$workingRow = 1
for($i = 0; $i -lt $pips.Count; $i++){
    $pip = $pips[$i]

    Write-Output "Exporting $($pip.Name)"

    $templatePackage.Workbook.Worksheets["pip"].Cells["A1:G${pipHeight}"].Copy($pipWs.Cells["A${workingRow}:G$($workingRow + $pipHeight)"])

    Set-ExcelRange -Worksheet $pipWs -Range "A$($workingRow)" -Value $pip.Name -Bold
    Add-ExcelName -Range $pipWs.Cells["A$($workingRow)"] -RangeName "pip_$($pip.Name)" -WarningAction SilentlyContinue
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 1)" -Value $pip.ResourceGroupName
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 2)" -Value $pip.Name
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 3)" -Value $pip.Location
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 4)" -Value $pip.PublicIpAllocationMethod
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 5)" -Value $pip.IpAddress
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 6)" -Value $pip.PublicIpAddressVersion
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 7)" -Value $pip.IdleTimeoutInMinutes
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 9)" -Value $pip.DnsSettings.DomainNameLabel
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 10)" -Value $pip.DnsSettings.Fqdn
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 11)" -Value $pip.DnsSettings.ReverseFqdn
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 14)" -Value $pip.Sku.Name
    Set-ExcelRange -Worksheet $pipWs -Range "G$($workingRow + 15)" -Value $pip.Sku.Tier
    $workingRow += $pipHeight + 1
}

Close-ExcelPackage $templatePackage
Close-ExcelPackage $excelPackage -Show

