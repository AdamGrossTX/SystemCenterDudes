<#

    .SYNOPSIS

    Find all collections with a RefreshType of 1 (Manual) that have a recurring schedule set and update the schedule to non-recurring (None).

    .NOTES

        Author: Adam Gross

        Twitter: @AdamGrossTX

        Website: https://www.asquaredozen.com
        
        NOTE - The script only handles Device collections. It should be modified for User collections if you need it. May update it at some point.

    .LINK
        Originally posted on http://www.SystemCenterDudes.com
        
        GitHub Repo http://www.github.com/AdamGrossTX

    .HISTORY

        1.0 - Original
#>

[cmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]
    $SiteCode,
    
    [Parameter(Mandatory=$true)]
    [string]
    $ProviderMachineName
)

#Connect to ConfigMgr
$initParams = @{}
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

Set-Location "$($SiteCode):\" @initParams

#######################################################

#Set New Blank Non-Recurring Schedule
$Schedule = New-CMSchedule -Start "01/01/2019 12:00 AM" -DurationInterval Minutes -DurationCount 0 -IsUtc:$False -Nonrecurring

#Get All Collections
$AllCollections = Get-CMDeviceCollection
Write-Host "Total Collections Count is: $($AllCollections.Count)"

#Filter to TargetCollections based on RefreshType of 1 which is Manual
$ManualRefreshCollections = $AllCollections | Where-Object RefreshType -eq 1
Write-Host "Total Collections with RefreshType of 1 is: $($ManualRefreshCollections.Count)"

#Get Collections with a RefreshSchedule that is recurring.
$RecurringCollections = $ManualRefreshCollections | Where-Object {$_.RefreshSchedule.SmsProviderObjectPath -ne "SMS_ST_NonRecurring"}
Write-Host "Total Collections with RefreshType of 1 and RefreshSchedule of Recurring: $($RecurringCollections.Count)"

$Count = 0
#Loop through each RecurringCollection and update the schedule to be non-recurring 
ForEach($Collection in $RecurringCollections)
{
    $Count ++
    Write-Host "#############################"
    Write-Host "Processing Record $($Count) of $($RecurringCollections.Count): $($Collection.Name)"
    $Collection | Set-CMDeviceCollection -RefreshSchedule $Schedule
    Write-Host "Updated: $($Collection.Name)"
}
