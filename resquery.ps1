<#
   .SYNOPSIS
   resquery
   
   .DESCRIPTION
   Azure Resource Query
   
   .PARAMETER SubscriptionId
   The Azure Subscription Id
   
   .EXAMPLE
   .\resquery.ps1 -SubscriptionId 00000000-0000-0000-0000-000000000000
   
   .NOTES
   General notes
   Copyright 2018 (c) MACROmantic
   Written by: christopher landry <macromantic (at) outlook.com>
   Version: 0.1.0
   Date: 10-november-2018
#>

param(
    [Parameter(Mandatory)]
    [string]$SubscriptionId,
    [Parameter(Mandatory=$false)]
    [string]$FilePath = ".\azure_resources_$(Get-Date -UFormat '%d_%m_%Y_%H_%M').xlsx"
)
$ErrorActionPreference = "stop"

function ConnectionCheck() {
    try {
        Get-AzureRmSubscription -SubscriptionId $SubscriptionId -WarningAction stop | Out-Null
        return
    } catch {
        Write-Warning "Not logged into Azure"
    }

    try {
        Connect-AzureRmAccount
        Set-AzureRmContext -Subscription $SubscriptionId 
    } catch {
        Write-Error "Error: Not logged in to $SubscriptionId"
    }
}

function GetResources() {
    $outputTypes = @{}

    Write-Host "Querying Azure Subscription $SubscriptionId"

    $resources = Get-AzureRmResource
    foreach ($resource in $resources) {
        $resourceType = $($resource.ResourceType.split("/"))[1]

        if (-not($outputTypes[$resourceType])) {
            $outputTypes[$resourceType] = @()
        }

        if ($resourceType -eq "disks") {
            $diskResource = Get-AzureRmDisk `
                -DiskName $resource.Name `
                -ResourceGroupName $resource.ResourceGroupName

            # More properties at:
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.azure.management.compute.models.disk?view=azure-dotnet
            $outputTypes[$resourceType] += @{
                Id = $resource.ResourceId;
                Name = $resource.Name;
                ResourceGroup = $resource.ResourceGroupName;
                Location = $resource.Location;
                DiskSizeGB = $diskResource.DiskSizeGB;
                OsType = $diskResource.OsType;
                Sku = $diskResource.Sku.Name;
            }
            
        } elseif ($resourceType -eq "virtualMachines") {
            $vmResource = Get-AzureRmVM `
                -Name $resource.Name `
                -ResourceGroupName $resource.ResourceGroupName

            # More properties at:
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.azure.management.compute.models.virtualmachine?view=azure-dotnet
            $outputTypes[$resourceType] += @{
                Id = $resource.ResourceId;
                Name = $resource.Name;
                ResourceGroup = $resource.ResourceGroupName;
                Location = $resource.Location;
                VMSize = $vmResource.HardwareProfile.VmSize;
                OsType = $vmResource.StorageProfile.OsDisk.OsType;
            }

        } elseif ($resourceType -eq "virtualNetworks") {
            $vnetResource = Get-AzureRmVirtualNetwork `
                -Name $resource.Name `
                -ResourceGroupName $resource.ResourceGroupName 3> $null

            $addressSpace = ConvertFrom-Json -InputObject $vnetResource.AddressSpace.AddressPrefixesText
       
            # More properties at:
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.azure.commands.network.models.psvirtualnetwork?view=azurerm-ps
            $outputTypes[$resourceType] += @{
                Id = $resource.ResourceId;
                Name = $resource.Name;
                ResourceGroup = $resource.ResourceGroupName;
                Location = $resource.Location;
                AddressSpace = "$addressSpace";
            }

        } elseif ($resourceType -eq "networkSecurityGroups") {
            $nsgResource = Get-AzureRmNetworkSecurityGroup `
                -Name $resource.Name `
                -ResourceGroupName $resource.ResourceGroupName

            $securityRules = ConvertFrom-Json -InputObject $nsgResource.SecurityRulesText;
            $rulesText = ""
            foreach ($rules in $securityRules) {
                $rulesText += " Rule: $($rules.Name) Protocol: $($rules.Protocol) Src: $($rules.SourcePortRange) Dest: $($rules.DestinationPortRange) Direction: $($rules.Direction) Access: $($rules.Access),"
            }
            if ($rulesText.Length -gt 0) {
                $rulesText = $rulesText.Substring(0, $rulesText.Length-1)
            }

            # More properties at:
            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.azure.commands.network.models.psnetworksecuritygroup?view=azurerm-ps 
            $outputTypes[$resourceType] += @{
                Id = $resource.ResourceId;
                Name = $resource.Name;
                ResourceGroup = $resource.ResourceGroupName;
                Location = $resource.Location;
                SecurityRules = "$rulesText";
            }
        } else {
            $outputTypes[$resourceType] += @{
                Id = $resource.ResourceId;
                Name = $resource.Name;
                ResourceGroup = $resource.ResourceGroupName;
                Location = $resource.Location;
            }
        }
    }

    return $outputTypes
}

function ExportToExcel() {
    param(
        [Parameter(Mandatory)]
        $ResourceHash
    )

    Write-Host "Exporting results to $FilePath"

    $excelapp = New-Object -ComObject Excel.Application
    $excelapp.visible = $false

    $workbook = $excelapp.workbooks.add()

    $sheetCount = 0
    foreach ($key in $ResourceHash.Keys) { 
        $sheet = $null
        if ($sheetCount -eq 0) {   
            $sheet = $workbook.sheets | Where-Object { $_.name -eq "Sheet1" }
        } elseif ($sheetCount -gt 0) {
            $sheet = $workbook.sheets.add()
        }
        
        $sheet.name = "$key"

        if ($key -eq "disks") {
            $rowCount = 1
            $sheet.range("A$($rowCount):A$($rowCount)").cells = "Name"
            $sheet.range("B$($rowCount):B$($rowCount)").cells = "Resource Group"
            $sheet.range("C$($rowCount):C$($rowCount)").cells = "Location"
            $sheet.range("D$($rowCount):D$($rowCount)").cells = "DiskSizeGB"
            $sheet.range("E$($rowCount):E$($rowCount)").cells = "OsType"
            $sheet.range("F$($rowCount):F$($rowCount)").cells = "Sku"
            
            foreach ($resource in $ResourceHash[$key]) {
                $rowCount++

                $sheet.range("A$($rowCount):A$($rowCount)").cells = "$($resource.Name)"
                $sheet.range("B$($rowCount):B$($rowCount)").cells = "$($resource.ResourceGroup)"
                $sheet.range("C$($rowCount):C$($rowCount)").cells = "$($resource.Location)"
                $sheet.range("D$($rowCount):D$($rowCount)").cells = "$($resource.DiskSizeGB)"
                $sheet.range("E$($rowCount):E$($rowCount)").cells = "$($resource.OsType)"
                $sheet.range("F$($rowCount):F$($rowCount)").cells = "$($resource.Sku)"
            }   
            $tblObj = $sheet.ListObjects.Add(1, $sheet.range("A1:F$($rowCount)"),"", 1)
            $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $tblObj.Name = "$key"
        } elseif ($key -eq "virtualMachines") {
            $rowCount = 1
            $sheet.range("A$($rowCount):A$($rowCount)").cells = "Name"
            $sheet.range("B$($rowCount):B$($rowCount)").cells = "Resource Group"
            $sheet.range("C$($rowCount):C$($rowCount)").cells = "Location"
            $sheet.range("D$($rowCount):D$($rowCount)").cells = "VMSize"
            $sheet.range("E$($rowCount):E$($rowCount)").cells = "OSType"
            
            foreach ($resource in $ResourceHash[$key]) {
                $rowCount++

                $sheet.range("A$($rowCount):A$($rowCount)").cells = "$($resource.Name)"
                $sheet.range("B$($rowCount):B$($rowCount)").cells = "$($resource.ResourceGroup)"
                $sheet.range("C$($rowCount):C$($rowCount)").cells = "$($resource.Location)"
                $sheet.range("D$($rowCount):D$($rowCount)").cells = "$($resource.VMSize)"
                $sheet.range("E$($rowCount):E$($rowCount)").cells = "$($resource.OSType)"
            }   
            $tblObj = $sheet.ListObjects.Add(1, $sheet.range("A1:E$($rowCount)"),"", 1)
            $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $tblObj.Name = "$key"
        
        } elseif ($key -eq "virtualNetworks") {
            $rowCount = 1
            $sheet.range("A$($rowCount):A$($rowCount)").cells = "Name"
            $sheet.range("B$($rowCount):B$($rowCount)").cells = "Resource Group"
            $sheet.range("C$($rowCount):C$($rowCount)").cells = "Location"
            $sheet.range("D$($rowCount):D$($rowCount)").cells = "AddressSpace"
            
            foreach ($resource in $ResourceHash[$key]) {
                $rowCount++

                $sheet.range("A$($rowCount):A$($rowCount)").cells = "$($resource.Name)"
                $sheet.range("B$($rowCount):B$($rowCount)").cells = "$($resource.ResourceGroup)"
                $sheet.range("C$($rowCount):C$($rowCount)").cells = "$($resource.Location)"
                $sheet.range("D$($rowCount):D$($rowCount)").cells = "$($resource.AddressSpace)"
            }   
            $tblObj = $sheet.ListObjects.Add(1, $sheet.range("A1:D$($rowCount)"),"", 1)
            $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $tblObj.Name = "$key"
        
        } elseif ($key -eq "networkSecurityGroups") {
            $rowCount = 1
            $sheet.range("A$($rowCount):A$($rowCount)").cells = "Name"
            $sheet.range("B$($rowCount):B$($rowCount)").cells = "Resource Group"
            $sheet.range("C$($rowCount):C$($rowCount)").cells = "Location"
            $sheet.range("D$($rowCount):D$($rowCount)").cells = "SecurityRules"

            foreach ($resource in $ResourceHash[$key]) {
                $rowCount++

                $sheet.range("A$($rowCount):A$($rowCount)").cells = "$($resource.Name)"
                $sheet.range("B$($rowCount):B$($rowCount)").cells = "$($resource.ResourceGroup)"
                $sheet.range("C$($rowCount):C$($rowCount)").cells = "$($resource.Location)"
                $sheet.range("D$($rowCount):D$($rowCount)").cells = "$($resource.SecurityRules)"
                $sheet.range("D$($rowCount):D$($rowCount)").EntireColumn.AutoFit() | Out-Null
            }   
            $tblObj = $sheet.ListObjects.Add(1, $sheet.range("A1:D$($rowCount)"),"", 1)
            $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $tblObj.Name = "$key"
        
        } else {
            $rowCount = 1
            $sheet.range("A$($rowCount):A$($rowCount)").cells = "Name"
            $sheet.range("B$($rowCount):B$($rowCount)").cells = "Resource Group"
            $sheet.range("C$($rowCount):C$($rowCount)").cells = "Location"
            
            foreach ($resource in $ResourceHash[$key]) {
                $rowCount++

                $sheet.range("A$($rowCount):A$($rowCount)").cells = "$($resource.Name)"
                $sheet.range("B$($rowCount):B$($rowCount)").cells = "$($resource.ResourceGroup)"
                $sheet.range("C$($rowCount):C$($rowCount)").cells = "$($resource.Location)"
            }  
            $tblObj = $sheet.ListObjects.Add(1, $sheet.range("A1:C$($rowCount)"),"", 1)
            $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $tblObj.Name = "$key"
        }
        $sheetCount++
    }
    $workbook.saveas($FilePath)
    $excelapp.quit()
}

# Main #############################################

ConnectionCheck
$resourceHash = GetResources
ExportToExcel -ResourceHash $resourceHash