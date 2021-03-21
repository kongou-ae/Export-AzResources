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
$storageAccounts = Get-AzStorageAccount
$rsvs = Get-AzRecoveryServicesVault
for($i = 0; $i -lt $rsvs.Count; $i++){
    $vault = $rsvs[$i]
    $rsvProperty = Get-AzRecoveryServicesVaultProperty -VaultId $vault.ID
    $rsvs[$i] | Add-Member StorageModelType $rsvProperty.StorageModelType
    $rsvs[$i] | Add-Member StorageType $rsvProperty.StorageType
    $rsvs[$i] | Add-Member StorageTypeState $rsvProperty.StorageTypeState
    $rsvs[$i] | Add-Member EnhancedSecurityState $rsvProperty.EnhancedSecurityState
    $rsvs[$i] | Add-Member SoftDeleteFeatureState $rsvProperty.SoftDeleteFeatureState
    $rsvs[$i] | Add-Member encryptionProperties $rsvProperty.encryptionProperties

    $rsvbackupProperty = Get-AzRecoveryServicesBackupProperties -Vault $vault
    $rsvs[$i] | Add-Member BackupStorageRedundancy $rsvbackupProperty.BackupStorageRedundancy
    $rsvs[$i] | Add-Member CrossRegionRestore $rsvbackupProperty.CrossRegionRestore
}



#-------------------------------------------------------------------------------------------
# Initialize
#-------------------------------------------------------------------------------------------
# Load excel files.

$date = Get-Date -Format yyyyMMdd-hhmm
$fileName = "azReport-${date}.xlsx"
$templateFileName = "$PSScriptRoot/azReportTemplate.xlsx"
$temptemplateFileName = "$PSScriptRoot/azReportTemplateTemp.xlsx"
$fileName = "$PSScriptRoot/$fileName"

if( Test-Path $fileName ){
    Remove-Item $fileName
}

if( Test-Path $temptemplateFileName ){
    Remove-Item $temptemplateFileName   
}

Copy-Item $templateFileName $temptemplateFileName

# Creat default sheets
Export-Excel $fileName -WorksheetName "README" 
Export-Excel $fileName -WorksheetName "SUMMARY"

# Load default sheets
$excelPackage = Open-ExcelPackage -Path $fileName
$templatePackage = Open-ExcelPackage -Path $temptemplateFileName

#-------------------------------------------------------------------------------------------
# Create README
#-------------------------------------------------------------------------------------------
$readmeMsg = @(
    "",
    "Thanks for using Export-AzResources.",
    "If you find any issue or any request, please open an issue to https://github.com/kongou-ae/Export-AzResources.",
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

function New-Summary {

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

    $row += 1

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

    $row += 1

    # create the summary of storage account
    Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "Storage account" -FontSize 12 -Bold
    $row ++
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "PrimaryLocation" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "SkuName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "Kind" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    $row ++

    for($i=0;$i -lt $storageAccounts.Count; $i++){
        $storageAccount = $storageAccounts[$i]
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $storageAccount.StorageAccountName -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $storageAccount.ResourceGroupName -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $storageAccount.PrimaryLocation -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value $storageAccount.Sku.Name -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value $storageAccount.Kind -BorderAround Thin

        $formular = '=HYPERLINK("#storageAccounts!storageAccount_' + ($($storageAccount.StorageAccountName) -replace "-","_") + '","Link")'
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,7].Address -Formula $formular -BorderAround Thin
        $row++
    }

    $row += 1

    # create the summary of Recovery Service Vault
    Set-ExcelRange -Worksheet $summaryWs -Range "A${row}" -Value "Recovery Service Vault" -FontSize 12 -Bold
    $row ++
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value "`#" -Width 5 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value "Name" -Width 40 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value "ResourceGroupName" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value "Location" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value "BackupStorageRedundancy" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Value "Detail" -Width 20 -BackgroundColor $bgc -BorderAround Thin
    $row ++

    for($i=0;$i -lt $rsvs.Count; $i++){
        $rsv = $rsvs[$i]
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,1].Address -Value ($row -4) -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,2].Address -Value $rsv.Name -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,3].Address -Value $rsv.ResourceGroupName -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,4].Address -Value $rsv.Location -BorderAround Thin
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,5].Address -Value $rsv.BackupStorageRedundancy -BorderAround Thin

        $formular = '=HYPERLINK("#RecoveryServiceVault!rsv_' + ($($rsv.Name) -replace "-","_") + '","Link")'
        Set-ExcelRange -Worksheet $summaryWs -Range $summaryWs.Cells[$row,6].Address -Formula $formular -BorderAround Thin
        $row++
    }
}

function New-VmDetails() {

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
    $vmLinuxHeight = 61
    $workingRow = 1 # エクスポート中のリソースのスタート行
    for($i = 0; $i -lt $vms.Count; $i++){
        $vm = $vms[$i]

        Write-Output "Exporting $($vm.Name)"
        if ( $vm.StorageProfile.OsDisk.OsType -eq "Windows" ){
            $templatePackage.Workbook.Worksheets["VirtualMachine_windows"].Cells["A1:G60"].Copy($ws.Cells["A${workingRow}:G$($workingRow + $vmWinHeight)"])          
        }

        if ( $vm.StorageProfile.OsDisk.OsType -eq "Linux" ){
            $templatePackage.Workbook.Worksheets["VirtualMachine_linux"].Cells["A1:G61"].Copy($ws.Cells["A${workingRow}:G$($workingRow + $vmLinuxHeight)"])          
        }

        Set-ExcelRange -Worksheet $ws -Range "A${workingRow}" -Value $vm.Name -Bold; $workingRow++
        Add-ExcelName -Range $ws.Cells["A${workingRow}"] -RangeName "vm_$($vm.Name)" -WarningAction SilentlyContinue
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Name; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Location; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.LicenseType; $workingRow += 3
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.DiagnosticsProfile.BootDiagnostics.Enabled; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.DiagnosticsProfile.BootDiagnostics.StorageUri; $workingRow+=2
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.HardwareProfile.VmSize; $workingRow+=4
        # ToDo: Support multiple nics
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.NetworkProfile.NetworkInterfaces[0].Primary; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.NetworkProfile.NetworkInterfaces[0].id; $workingRow+=2
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.ComputerName; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.AdminUsername; $workingRow+=3

        if ( $vm.StorageProfile.OsDisk.OsType -eq "Windows" ){
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.ProvisionVMAgent; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.EnableAutomaticUpdates; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.TimeZone; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.AdditionalUnattendContent; $workingRow+=2
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.PatchSettings.PatchMode; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.PatchSettings.EnableHotpatching; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.WindowsConfiguration.WinRM; $workingRow++
        }    
        if ( $vm.StorageProfile.OsDisk.OsType -eq "Linux" ){
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.LinuxConfiguration.DisablePasswordAuthentication; $workingRow+=4
            if ($vm.OSProfile.LinuxConfiguration.Ssh -ne $null){
                Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.LinuxConfiguration.Ssh.PublicKeys[0].Path; $workingRow++
                Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.LinuxConfiguration.Ssh.PublicKeys[0].KeyData; $workingRow++    
            } else {
                $workingRow+=2
            }
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.LinuxConfiguration.ProvisionVMAgent; $workingRow+=2
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.LinuxConfiguration.PatchSettings.PatchMode; $workingRow++
        }

        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.Secrets; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.OSProfile.AllowExtensionOperations; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.BillingProfile; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Plan; $workingRow+=3
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.ImageReference.Publisher; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.ImageReference.Offer; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.ImageReference.Sku; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.ImageReference.Version; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.ImageReference.ExactVersion; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.ImageReference.id; $workingRow+=2
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.OsType; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.EncryptionSettings; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.Name; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.Caching; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.WriteAcceleratorEnabled; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.CreateOption; $workingRow+=2
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.ManagedDisk.StorageAccountType; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.ManagedDisk.DiskEncryptionSet; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.StorageProfile.OsDisk.ManagedDisk.id; $workingRow++
        
        $workingRow++
        for($j=0;$j -lt $vm.StorageProfile.DataDisks.Count;$j++ ){
            Write-Output "Adding $($vm.StorageProfile.DataDisks[$j].Name) to $($vm.Name)"
            $addedRowNumbers = 11 # 足される行数

            $ws.InsertRow($workingRow,$addedRowNumbers) 
            $templatePackage.Workbook.Worksheets["DataDisk"].Cells["A1:G${addedRowNumbers}"].Copy($ws.Cells["A${workingRow}:G$($workingRow + $addedRowNumbers -1)"])

            $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].Lun; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].Name; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].Caching; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].CreateOption; $workingRow+=2
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].ManagedDisk.StorageAccountType; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].ManagedDisk.DiskEncryptionSet; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].ManagedDisk.Id; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].ToBeDetached; $workingRow++
            Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value  $vm.StorageProfile.DataDisks[$j].DetachOption; $workingRow++
        }
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Identity; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Zones; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.ProximityPlacementGroup; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Host; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.EvictionPolicy; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.Priority; $workingRow++
        Set-ExcelRange -Worksheet $ws -Range "G${workingRow}" -Value $vm.HostGroup; $workingRow++
        $workingRow++

    }

}

function New-VnetDetails() {

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

        Set-ExcelRange -Worksheet $vnetWs -Range "A$($workingRow)" -Value $vnet.Name -Bold; $workingRow++
        Add-ExcelName -Range $vnetWs.Cells["A$($workingRow)"] -RangeName "vnet_$($vnet.Name)" -WarningAction SilentlyContinue;
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value $vnet.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value $vnet.Name; $workingRow++
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value $vnet.Location; $workingRow+=2
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value ($vnet.AddressSpace.AddressPrefixes -join ","); $workingRow+=2
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value ($vnet.DhcpOptions.DnsServers -join ","); $workingRow++

        $workingRow++
        for($j=0;$j -lt $vnet.Subnets.Count;$j++ ){
            Write-Output "Adding $($vnet.Subnets[$j].Name) to $($vnet.Name)"
            $addedRowNumbers = 15 # 足される行数
            $vnetWs.InsertRow($workingRow,$addedRowNumbers) 
            $templatePackage.Workbook.Worksheets["Subnet"].Cells["A1:G${addedRowNumbers}"].Copy($vnetWs.Cells["A${workingRow}:G$($workingRow + $addedRowNumbers -1)"])
            
            $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].Name; $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  ($vnet.Subnets[$j].AddressPrefix -join ","); $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].ServiceAssociationLinks; $workingRow+=2
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].NetworkSecurityGroup.Id; $workingRow+=2
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].RouteTable.DisableBgpRoutePropagation; $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].RouteTable.Id; $workingRow+=2
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].NatGateway.Id; $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].ServiceEndpoints; $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].ServiceEndpointPolicies; $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].PrivateLinkServiceNetworkPolicies; $workingRow++
            Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value  $vnet.Subnets[$j].Delegations; $workingRow++

        }
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value $vnet.VirtualNetworkPeerings; $workingRow+=1
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value $vnet.EnableDdosProtection; $workingRow+=1
        Set-ExcelRange -Worksheet $vnetWs -Range "G${workingRow}" -Value $vnet.DdosProtectionPlan; $workingRow+=1
        $workingRow+=1
    }
}


function New-DiskDetails() {

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

        Set-ExcelRange -Worksheet $diskWs -Range "A$($workingRow)" -Value $disk.Name -Bold; $workingRow++
        Add-ExcelName -Range $diskWs.Cells["A$($workingRow)"] -RangeName "disk_$($disk.Name)" -WarningAction SilentlyContinue
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.Name; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.Location; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.ManagedBy; $workingRow+=2
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.sku.Name; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.sku.Tier; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.Zone; $workingRow+=2
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.CreationData.CreateOption; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.CreationData.StorageAccountId; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.CreationData.ImageReference; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.CreationData.GalleryImageReference; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.DiskSizeGB; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.DiskState; $workingRow+=2
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.Encryption.DiskEncryptionSetId; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.Encryption.Type; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.ShareInfo; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.NetworkAccessPolicy; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.Tier; $workingRow++
        Set-ExcelRange -Worksheet $diskWs -Range "G${workingRow}" -Value $disk.BurstingEnabled; $workingRow++

        $workingRow++
    }
}


function New-NicDetails() {
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

        Set-ExcelRange -Worksheet $nicWs -Range "A$($workingRow)" -Value $nic.Name -Bold; $workingRow++
        Add-ExcelName -Range $nicWs.Cells["A$($workingRow)"] -RangeName "nic_$($nic.Name)" -WarningAction SilentlyContinue
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.Name; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.Location; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.VirtualMachine; $workingRow+=3
        # ToDo: Support multiple ip configuration
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].Name; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].PrivateIpAddress; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].PrivateIpAllocationMethod; $workingRow+=2
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].Subnet.Id; $workingRow+=2
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].PublicIpAddress.Id; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].PrivateIpAddressVersion; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].LoadBalancerBackendAddressPools; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].Primary; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].ApplicationGatewayBackendAddressPools; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.IpConfigurations[0].ApplicationSecurityGroups; $workingRow+=2
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.DnsSettings.DnsServers; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.DnsSettings.AppliedDnsServers; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.EnableIPForwarding; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.EnableAcceleratedNetworking; $workingRow+=2
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.NetworkSecurityGroup.Id; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.Primary; $workingRow++
        Set-ExcelRange -Worksheet $nicWs -Range "G${workingRow}" -Value $nic.MacAddress; $workingRow++

        $workingRow++
    }
}


function New-NsgDetails {

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
            Set-ExcelRange -Worksheet $nsgWs -Range "A${workingRow}" -Value $sortedNsgRules[$j].Direction -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "A${workingRow}:E${workingRow}" -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "F${workingRow}" -Value $sortedNsgRules[$j].Priority -BorderAround Thin -BackgroundColor white   
            Set-ExcelRange -Worksheet $nsgWs -Range "G${workingRow}" -Value $sortedNsgRules[$j].Name -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "H${workingRow}" -Value $sortedNsgRules[$j].Protocol -BorderAround Thin -BackgroundColor white   
            Set-ExcelRange -Worksheet $nsgWs -Range "I${workingRow}" -Value $sortedNsgRules[$j].SourceAddressPrefix -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "J${workingRow}" -Value $sortedNsgRules[$j].SourceApplicationSecurityGroups -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "K${workingRow}" -Value $sortedNsgRules[$j].SourcePortRange -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "L${workingRow}" -Value $sortedNsgRules[$j].DestinationAddressPrefix -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "M${workingRow}" -Value $sortedNsgRules[$j].DestinationApplicationSecurityGroups -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "N${workingRow}" -Value $sortedNsgRules[$j].DestinationPortRange -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "O${workingRow}" -Value $sortedNsgRules[$j].Access -BorderAround Thin -BackgroundColor white
            Set-ExcelRange -Worksheet $nsgWs -Range "P${workingRow}" -Value $sortedNsgRules[$j].Description -BorderAround Thin -BackgroundColor white
            $workingRow ++ 
        }

        $nsg.NetworkInterfaces | ForEach-Object {
            $workingRow ++ # 空の行の位置に移動
            $k = 0
            $nsgWs.InsertRow($workingRow,2) ## 空行とID行を追加
            $templatePackage.Workbook.Worksheets["Nsg1stArrayOnlyId"].Cells["A1:K2"].Copy($nsgWs.Cells["A${workingRow}:P$($workingRow + 1)"])
            $workingRow ++ # ID 行に移動
            Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow)" -Value $nsg.NetworkInterfaces[$k].Id -BorderAround Thin -BackgroundColor white
            $k++
        }

        $workingRow ++ 
        $nsg.Subnets | ForEach-Object {
            $workingRow ++ # 空の行の位置に移動
            $k = 0
            $nsgWs.InsertRow($workingRow,2)
            $templatePackage.Workbook.Worksheets["Nsg1stArrayOnlyId"].Cells["A1:K2"].Copy($nsgWs.Cells["A${workingRow}:P$($workingRow + 1)"])
            $workingRow ++ 
            Set-ExcelRange -Worksheet $nsgWs -Range "G$($workingRow)" -Value $nsg.Subnets[$k].Id -BorderAround Thin -BackgroundColor white
            $k++

        }
        Set-ExcelRange -Worksheet $nsgWs -Range "A$($workingRow):G$($workingRow)" -BorderBottom Thin

        $workingRow += 2 # NetworkInterfaces,Subnets, space の3つ
    }
}

function New-PipDetails {
        
    #-------------------------------------------------------------------------------------------
    # Create pip
    #-------------------------------------------------------------------------------------------

    if ( $pips -ne $Null ){
        Write-Output "Adding the new worksheet for Network Interface"
        $pipWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "PublicIps"
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

        Set-ExcelRange -Worksheet $pipWs -Range "A$($workingRow)" -Value $pip.Name -Bold; $workingRow++
        Add-ExcelName -Range $pipWs.Cells["A$($workingRow)"] -RangeName "pip_$($pip.Name)" -WarningAction SilentlyContinue
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.Name; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.Location; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.PublicIpAllocationMethod; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.IpAddress; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.PublicIpAddressVersion; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.IdleTimeoutInMinutes; $workingRow+=2
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.DnsSettings.DomainNameLabel; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.DnsSettings.Fqdn; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.DnsSettings.ReverseFqdn; $workingRow++
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.Zone; $workingRow+=2
        Set-ExcelRange -Worksheet $pipWs -Range "G${workingRow}" -Value $pip.Sku.Name; $workingRow++
        $workingRow++
    }
}

function New-StorageAccountDetails {

    #-------------------------------------------------------------------------------------------
    # Create storage account
    #-------------------------------------------------------------------------------------------

    if ( $storageAccounts -ne $Null ){
        Write-Output "Adding the new worksheet for storage account"
        $storageAccountWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "StorageAccounts"
        $shortCols = @("A","B","C","D","E")
        foreach ($shortCol in $shortCols) {
            Set-ExcelRange -Worksheet $storageAccountWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
        }
        $shortCols = @("F","G","H","I","J","K")
        foreach ($shortCol in $shortCols) {
            Set-ExcelRange -Worksheet $storageAccountWs -Range "${shortCol}:${shortCol}" -Width (20).ToString()
        }
    }

    $storageAccountHeight = 45
    $workingRow = 1
    for($i = 0; $i -lt $storageAccounts.Count; $i++){
        $storageAccount = $storageAccounts[$i]

        Write-Output "Exporting $($storageAccount.StorageAccountName)"

        $templatePackage.Workbook.Worksheets["StorageAccount"].Cells["A1:K${storageAccountHeight}"].Copy($storageAccountWs.Cells["A${workingRow}:K$($workingRow + $storageAccountHeight)"])

        Set-ExcelRange -Worksheet $storageAccountWs -Range "A$($workingRow)" -Value $storageAccount.StorageAccountName -Bold; $workingRow++
        Add-ExcelName -Range $storageAccountWs.Cells["A$($workingRow)"] -RangeName "storageAccount_$($storageAccount.StorageAccountName)" -WarningAction SilentlyContinue
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.StorageAccountName; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.Location; $workingRow+=2
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.Sku.Name; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.Sku.Tier; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.Kind; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.AccessTier; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.CustomDomain; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.Identity; $workingRow+=2
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.Blob; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.Queue; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.Table; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.File; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.Web; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.Dfs; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.MicrosoftEndpoints; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryEndpoints.InternetEndpoints; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.PrimaryLocation; $workingRow+=2
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.Blob; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.Queue; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.Table; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.File; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.Web; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.Dfs; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.MicrosoftEndpoints; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryEndpoints.InternetEndpoints; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.SecondaryLocation; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.EnableHttpsTrafficOnly; $workingRow+=2
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.AzureFilesIdentityBasedAuth.DirectoryServiceOptions; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.AzureFilesIdentityBasedAuth.ActiveDirectoryProperties; $workingRow+=2
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.NetworkRuleSet.Bypass; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.NetworkRuleSet.DefaultAction; $workingRow++

        if ($storageAccounts.NetworkRuleSet.IpRules.Count -ne 0 ){
            $workingRow++
            $storageAccount.NetworkRuleSet.IpRules | ForEach-Object {
                Write-Host "Insert iprule to $($storageAccount.StorageAccountName)"
                $ipRule = $_
                $storageAccountWs.InsertRow($workingRow,1)
                $templatePackage.Workbook.Worksheets["StorageAccountAddOn"].Cells["A3:K3"].Copy($storageAccountWs.Cells["A${workingRow}:K${workingRow}"])
                Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $IpRule.Action
                Set-ExcelRange -Worksheet $storageAccountWs -Range "H${workingRow}" -Value $IpRule.IPAddressOrRange
                $workingRow++
            }    
        } else {
            $workingRow++
        }

        if ($storageAccounts.NetworkRuleSet.VirtualNetworkRules.Count -ne 0 ){
            $workingRow++
            $storageAccount.NetworkRuleSet.VirtualNetworkRules | ForEach-Object {
                Write-Host "Insert iprule to $($storageAccount.StorageAccountName)"
                $VirtualNetworkRule = $_
                $storageAccountWs.InsertRow($workingRow,1)
                $templatePackage.Workbook.Worksheets["StorageAccountAddOn"].Cells["A6:K6"].Copy($storageAccountWs.Cells["A${workingRow}:K${workingRow}"])
                Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $VirtualNetworkRule.Action
                Set-ExcelRange -Worksheet $storageAccountWs -Range "H${workingRow}" -Value $VirtualNetworkRule.VirtualNetworkResourceId
                $workingRow++
            }    
        } else {
            $workingRow++
        }

        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.NetworkRuleSet.ResourceAccessRules; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.RoutingPreference; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.AllowBlobPublicAccess; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.EnableNfsV3; $workingRow++
        Set-ExcelRange -Worksheet $storageAccountWs -Range "G${workingRow}" -Value $storageAccount.AllowSharedKeyAccess; $workingRow++
       
        $workingRow++
    }

}

function New-RecoveryServiceVaultDetails {

    #-------------------------------------------------------------------------------------------
    # Create recovery service vault
    #-------------------------------------------------------------------------------------------

    if ( $rsvs -ne $Null ){
        Write-Output "Adding the new worksheet for recovery service vault"
        $rsvWs = Add-Worksheet -ExcelPackage $excelPackage -WorksheetName "RecoveryServiceVaults"
        $shortCols = @("A","B","C","D","E")
        foreach ($shortCol in $shortCols) {
            Set-ExcelRange -Worksheet $rsvWs -Range "${shortCol}:${shortCol}" -Width (20/7).ToString()
        }
        Set-ExcelRange -Worksheet $rsvWs -Range "F:F" -Width (20).ToString()
        Set-ExcelRange -Worksheet $rsvWs -Range "G:G" -Width (100).ToString()
    }

    $rsvHeight = 20
    $workingRow = 1
    for($i = 0; $i -lt $rsvs.Count; $i++){
        $rsv = $rsvs[$i]

        Write-Output "Exporting $($rsv.Name)"

        $templatePackage.Workbook.Worksheets["RecoveryServiceVault"].Cells["A1:G${rsvHeight}"].Copy($rsvWs.Cells["A${workingRow}:G$($workingRow + $rsvHeight)"])

        Set-ExcelRange -Worksheet $rsvWs -Range "A$($workingRow)" -Value $rsv.Name -Bold; $workingRow++
        Add-ExcelName -Range $rsvWs.Cells["A$($workingRow)"] -RangeName "rsv$($rsv.Name)" -WarningAction SilentlyContinue
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.ResourceGroupName; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.Name; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.Location; $workingRow+=2
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.PrivateEndpointStateForBackup; $workingRow+=2
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.PrivateEndpointStateForSiteRecovery; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.StorageModelType; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.StorageType; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.StorageTypeState; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.EnhancedSecurityState; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.SoftDeleteFeatureState; $workingRow+=2
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.EncryptionAtRestType; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.KeyUri; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.InfrastructureEncryptionState; $workingRow+=2
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.BackupStorageRedundancy; $workingRow++
        Set-ExcelRange -Worksheet $rsvWs -Range "G${workingRow}" -Value $rsv.CrossRegionRestore; $workingRow++

        $workingRow++
    }

}

New-Summary
New-VmDetails
New-VnetDetails
New-DiskDetails
New-NicDetails
New-NsgDetails
New-PipDetails
New-StorageAccountDetails
New-RecoveryServiceVaultDetails

Close-ExcelPackage $templatePackage
Close-ExcelPackage $excelPackage -Show

Remove-Item $temptemplateFileName
Write-Output "Complete. Generates ${fileName}"

