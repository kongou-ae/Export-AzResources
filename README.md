# Export-AzResources

Convert your Azure resources to an Excel parameter sheet automatically.

## Sample


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

- Virutal Machine
- Virtual Network
- Disk
- Network interface
- Network Security Group
- Public Ip Address

## Lisence

MIT
