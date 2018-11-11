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
   Version: 0.0.3
   Date: 10-november-2018
#>

param(
    [Parameter(Mandatory)]
    [string]$SubscriptionId,
    [Parameter(Mandatory=$false)]
    [string]$FilePath = ".\requery.xlsx"
)
$ErrorActionPreference = "stop"

function ConnectionCheck() {
    try {
        Get-AzureRmSubscription -SubscriptionId $SubscriptionId | Out-Null
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

    $resources = Get-AzureRmResource
    foreach ($resource in $resources) {
        $resourceType = $($resource.ResourceType.split("/"))[1]

        if (-not($outputTypes[$resourceType])) {
            $outputTypes[$resourceType] = @()
        }
        $outputTypes[$resourceType] += @{
            Id = $resource.ResourceId;
            Name = $resource.Name;
            ResourceGroup = $resource.ResourceGroupName;
            Location = $resource.Location;
        }
    }

    return $outputTypes
}

function ExportToExcel() {
    param(
        [Parameter(Mandatory)]
        $ResourceHash
    )

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
        
        $sheetCount++
    }
    $workbook.saveas($FilePath)
    $excelapp.quit()
}

# Main #############################################

ConnectionCheck
$resourceHash = GetResources
ExportToExcel -ResourceHash $resourceHash