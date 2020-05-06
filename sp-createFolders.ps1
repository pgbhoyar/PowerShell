<#
.SYNOPSIS
    Script to create folder structure inside Microsoft's Teams SharePoint site using input from a .csv file
.DESCRIPTION
    This script comes handy if you would like to migrate the documents from subsites to Microsoft Teams.
    This script uses SharePoint PnP PowerShell Module for SharePoint Online
    https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

.EXAMPLE
	.\sp-CreateFolders.ps1 
.PARAMETER
	.
.NOTES
 *************************************************************************
 *  Copyright (C) 2020-2021 Prashant G Bhoyar and Withum Digital, LLC.
 *  All Rights Reserved.
 *
 * NOTICE:  All information contained herein is, and remains
 * the property of Prashant G Bhoyar and Withum Digital, LLC.
 * The intellectual and technical concepts contained
 * herein are proprietary to Prashant G Bhoyar and Withum Digital
 * and may be covered by U.S. and Foreign Patents, patents in
 * process, and are protected by trade secret or copyright law.
 * Dissemination of this information or reproduction of this material
 * is strictly forbidden unless prior written permission is obtained
 * from Prashant G Bhoyar and Withum Digital, LLC.
 *************************************************************************
#>

#Install the SharePoint PnP PowerShell Module https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps
Import-Module SharePointPnPPowerShellOnline
Write-Output "$([System.DateTime]::Now), Execution Started"
Write-Host "$([System.DateTime]::Now), Execution Started"

$csvFile = "C:\Users\pgbhoyar\Desktop\OneDrive - Prashant G Bhoyar\Work\Extranet\MigrationScript\folderstructure.csv";
Write-Output "$([System.DateTime]::Now), Before CSV Import"	
Write-Host "$([System.DateTime]::Now), Before CSV Import"	
$table = Import-Csv $csvFile -Delimiter ","
Write-Output "$([System.DateTime]::Now), after CSV Import"
Write-Host "$([System.DateTime]::Now), after CSV Import"
$count = 1;	
$credentials =  Get-Credential
Connect-PnPOnline -Url https://withumonline.sharepoint.com/sites/WD-LegacyExtranet/ -Credentials $credentials
                
foreach ($row in $table) {
    Write-Output "$([System.DateTime]::Now), processing row no : $count"
    Write-Host "$([System.DateTime]::Now), processing row no : $count"
    $count++;
    $siteUrl = [System.Uri] $($row."Siteaddress").Trim()
    $destinationFolderName = $($row."FolderName").Trim()
    # Typical Folder Structure in Microsoft Teams
    $folderPath = "Shared Documents/General/General/00_Legacy/" + $destinationFolderName
    
    try 
    {
        Resolve-PnPFolder -SiteRelativePath $folderPath;
    }
    catch 
    {
        $ErrorMessage = $_.Exception.Message
        Write-Output "$([System.DateTime]::Now), $_.Exception.Message";
        Write-Host "$([System.DateTime]::Now), $_.Exception.Message";
        Write-Output "$([System.DateTime]::Now), Error: occured while creating the folder $destinationFolderName";
        Write-Host "$([System.DateTime]::Now), Error: occured while creating the document library $destinationFolderName";
        #continue;
    }
}

Write-Output "$([System.DateTime]::Now), Execution Ended"
Write-Host "$([System.DateTime]::Now), Execution Ended"
