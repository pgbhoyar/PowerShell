<#
.SYNOPSIS
    Script to add owners to existing Flows ( Microsoft Power Automate)
.DESCRIPTION
    This script comes handy if you would like to add owners to existing Flows ( Microsoft Power Automate).
    This script uses PowerShell Modules for Power Apps and Power Automate
    https://docs.microsoft.com/en-us/power-platform/admin/powerapps-powershell

.EXAMPLE
	.\AddOwnerToFlows.ps1 
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


### To install the required PowerShell Modules
#Install-Module -Name Microsoft.PowerApps.Administration.PowerShell
#Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber
#Install-Module -Name AzureAD

###Email address of the account that needs to be added as owner in all Flows
$emailAddress = "demo@pgbhoyar.onmicrosoft.com";


###To Login with Flow Admin Account ( Right Now Global Admin Account) of the Tenant where we are running this script. 
Add-PowerAppsAccount 
Connect-AzureAD

$flowEnvironments = Get-FlowEnvironment

Write-Host "Flow Environment ID is " $flowEnvironments[0].EnvironmentName
$userID = Get-AzureADUser -ObjectID $emailAddress | Select-Object ObjectId

Write-Host "UserID is " $userID.objectId


$flows = Get-AdminFlow | Select-Object FlowName, DisplayName
foreach($flow in $flows){
    Write-Host "Adding Owner in the Flow "$flow.DisplayName
    Write-Host "Adding Owner in the Flow "$flow.FlowName
    try{
        ## Sample cmdlet
        #Set-AdminFlowOwnerRole -PrincipalType Group -PrincipalObjectId <Guid> -RoleName CanEdit -FlowName <Guid> -EnvironmentName Default-<Guid>
        
        ## If you have single environment
        #Set-AdminFlowOwnerRole -PrincipalType User -PrincipalObjectId $userID.objectId -RoleName CanEdit -FlowName $flow.FlowName -EnvironmentName $flowEnvironments.EnvironmentName
        
        ## If you have multiple environments
        Set-AdminFlowOwnerRole -PrincipalType User -PrincipalObjectId $userID.objectId -RoleName CanEdit -FlowName $flow.FlowName -EnvironmentName $flowEnvironments[0].EnvironmentName
        
        Write-Host "Added Owner" $emailAddress " to the Flow " $flow.DisplayName
        }
    Catch [System.Exception]{
       Write-Host $_.Exception.Message
   }
    
}