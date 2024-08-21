﻿# Including all functions
        # Get all PS1 files from the Functions folder
        $functionFiles = Get-ChildItem -Path "$PSScriptRoot\Functions\*.ps1"

        # Import each PS1 file
        foreach ($file in $functionFiles) {
            . $file.FullName
        }

# Main Menu
function M365cdeAAASetup(){
<#
.SYNOPSIS
Is starting the main menu for the Azure Automation Account Setup.

.DESCRIPTION
This function is starting the main menu for the Azure Automation Account Setup. It is the entry point for the module.
#>
    Clear-Host
    switch(Read-Host "

   ░▒█▀▄▀█░█▀▀█░▄▀▀▄░█▀▀░░░░█▀▄░▄▀▀▄░█▀▀▄░█▀▀░█░▒█░█░░▀█▀░█▀▀▄░█▀▀▄░▀█▀░░░░█▀▄░█▀▀
   ░▒█▒█▒█░░▒▀▄░█▄▄░░▀▀▄░▀▀░█░░░█░░█░█░▒█░▀▀▄░█░▒█░█░░░█░░█▄▄█░█░▒█░░█░░▄▄░█░█░█▀▀
   ░▒█░░▒█░█▄▄█░▀▄▄▀░▄▄▀░░░░▀▀▀░░▀▀░░▀░░▀░▀▀▀░░▀▀▀░▀▀░░▀░░▀░░▀░▀░░▀░░▀░░▀▀░▀▀░░▀▀▀
                                                                ©️2024 Dominik Gilgen


               ▂▃▅▆ █ Azure Automation Account Setup 0.1.0 █ ▆▅▃▂



This module helps you setting up and managing an Azure Automation Account.

--------------------------------------------------
1 Microsoft Graph Module
2 Microsoft Azure Module
--------------------------------------------------
3 Automation Account and Modules
4 Managed Identity and Permissions
--------------------------------------------------
5 Maester Test Framework (beta)
--------------------------------------------------

q Exit

Select"){
        q {break}
        1 {M365cdeGraphModule}
        2 {M365cdeAzModule}
        3 {if (Get-AzContext) { if ($AutomationAccountName) {M365cdeAutomationAccount} else {az_automation_set} } else { Write-Warning "No active Azure connection! Connect to Azure and return!"; Start-Sleep -Seconds 3; M365cdeAzModule}}
        4 {if (Get-MgContext) {M365cdeMIDgraph} else { Write-Warning "No active Microsoft Graph connection! Connect to Microsoft Graph and return!"; Start-Sleep -Seconds 3; M365cdeGraphModule}}
        5 {
            if ((Get-AzContext) -and (Get-MgContext)) {
                if ($AutomationAccountName) {M365cdeMaester}
                else {az_automation_set -breadcrumb "M365cdeMaester"}
            } else {
                # Check if there is an active Azure and Microsoft Graph connection
                if (-not (Get-AzContext)) {Write-Warning "No active Azure connection! Connect to Azure and return!"}
                if (-not (Get-MgContext)) {Write-Warning "No active Microsoft Graph connection! Connect to Microsoft Graph and return!"}
                Start-Sleep -Seconds 3; M365cdeAAASetup
            }
        }
        default {M365cdeAAASetup}
    }
}


Export-ModuleMember -Function M365cdeAAASetup,M365cdeGraphModule,M365cdeAzModule