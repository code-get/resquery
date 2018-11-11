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
   Version: 0.0.2
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

function GetResourceTypes() {
    $outputTypes = @{}
    $resourceTypes = Get-AzureRmResource | Select-Object ResourceType -Unique
    foreach ($resource in $resourceTypes) {
        $nameTokens = $resource.ResourceType.split("/")
        $outputTypes[$nameTokens[1]] = 0
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
    $sheet1 = $workbook.sheets | Where-Object { $_.name -eq "Sheet1" }

    $sheetCount = 0
    foreach ($resource in $ResourceHash.Keys) { 
        $sheet = $null
        if ($sheetCount -eq 0) {   
            $sheet = $sheet1
        } elseif ($sheetCount -gt 0) {
            $sheet = $workbook.sheets.add()
        }
        
        $sheet.name = "$resource"
        
        $sheetCount++
    }
    $workbook.saveas($FilePath)
    $excelapp.quit()
}

# Main #############################################

ConnectionCheck
$resourceHash = GetResourceTypes
ExportToExcel -ResourceHash $resourceHash