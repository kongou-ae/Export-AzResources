# Export-AzResources

Convert your Azure resources to an Excel parameter sheet automatically.

## Sample

[azReport_sample.xlsx](https://github.com/kongou-ae/Export-AzResources/raw/master/example/azReport_sample.xlsx)

## Usage

```
git clone https://github.com/kongou-ae/Export-AzResources
cd Export-AzResources
Install-Module ImportExcel -scope CurrentUser
Install-Module Az -scope CurrentUser
 .\Export-AzResources.ps1
```

## Requirement

- [dfinke/ImportExcel](https://github.com/dfinke/ImportExcel/)
- [Azure Powershell](https://docs.microsoft.com/ja-jp/powershell/azure/?view=azps-5.6.0)

## Supported resoruces

- Resource Group
- Virutal Machine
- Virtual Network
- Disk
- Network interface
- Network Security Group
- Public Ip Address
- Storage Account
- Recovery Service Vault

Notes: Every resorces may contain un-supported properties.

## Lisence

MIT
