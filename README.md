# resquery - Azure Resource Query
* Copyright 2018 (c) MACROmantic
* Written by: christopher landry <macromantic (at) outlook.com>
* Version: 0.1.2
* Initial Date: 10-november-2018

### Requirements (Host)
[Azure PowerShell v5.8.0](https://github.com/Azure/azure-powershell/releases/tag/v5.7.0-April2018)

### Run Instructions

```
  PS C:\Users\azadmin> .\resquery.ps1 -SubscriptionId 00000000-0000-0000-0000-000000000000
```
This will output a file at C:\Users\azadmin\Documents\azure_resource_query_date.xlsx

or with an optional file path
```
  PS C:\Users\azadmin> .\resquery.ps1 -SubscriptionId 00000000-0000-0000-0000-000000000000 -FilePath C:\azure_resource_file.xlsx
```

#### Version 0.0.1
* Initial Check-in

#### Version 0.0.2
* Added Azure Login Check
* Query Resource Types
* Export Resource Types as Tab in Excel

#### Version 0.0.3
* Row Data Populates

#### Version 0.0.4
* Fixed login check

#### Version 0.1.0
* Display Properties By Resource Type
* Added Basic Table Formatting

#### Version 0.1.1
* Added more VirtualNetwork columns

#### Version 0.1.2
* Added NSG Rule formatting