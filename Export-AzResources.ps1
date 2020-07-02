$ErrorActionPreference = "stop"


function Confirm-AzModule {

    $installedModules = Get-module -ListAvailable
    $importedModules = Get-module

    if ($null -eq ($installedModules | Where-Object {$_.Name -eq "Az"})){
        Write-Output "Install-Module Az -scope local"
        Install-Module Az -scope currentuser -force
    }    

    if ($null -eq ($importedModules | Where-Object {$_.Name -eq "Az.Resources"})){
        Write-Output "Import-Module Az.Resources -scope local"
        Import-Module Az.Resources -scope local -force
    }
}

function Clear-oldJson {
    $files = Get-ChildItem -Path ./
    $files = $files | Where-Object {$_.Name -like "*.json"} 
    if ( $null -ne $files ){
        Write-Output "Delete the following files"
        Write-Output $files
        $files | Remove-Item -Force
    }
}

Confirm-AzModule
Clear-oldJson

$resourceTypes = New-Object System.Collections.ArrayList

$resources = Get-AzResource -ExpandProperties
$resources | Select-Object ResourceType -Unique | ForEach-Object {
    $resourceTypes += $_.ResourceType
} 

$resourceTypes | ForEach-Object {
    $resourceType = $_
    $fileName = $resourceType -replace "/","." 
    $resources | Where-Object { $_.ResourceType -eq $resourceType } | ConvertTo-Json -Depth 100 | Out-File "$fileName.json" -Force
}



$fileName = $resourceTypes -replace "/","." 



