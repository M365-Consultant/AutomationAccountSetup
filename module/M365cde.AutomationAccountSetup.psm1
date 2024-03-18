#Including all functions
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


               ▂▃▅▆ █ Azure Automation Account Setup 0.0.2 █ ▆▅▃▂



This Script helps you to setting up an Automation Account.

--------------------------------------------------
1 Microsoft Graph Module
2 Microsoft Azure
3 Automation Account
4 Configure Managed Identity

--------------------------------------------------

q Exit

Select"){
        q {break}
        1 {M365cdeGraphModule}
        2 {M365cdeAzModule}
        3 {if (Get-AzContext) { if ($AutomationAccountName) {M365cdeAutomationAccount} else {az_automation_set} } else { Write-Warning "No active Azure connection! Connect to Azure and return!"; Start-Sleep -Seconds 3; M365cdeAzModule}}
        4 {if (Get-MgContext) {M365cdeMIDgraph} else { Write-Warning "No active Microsoft Graph connection! Connect to Microsoft Graph and return!"; Start-Sleep -Seconds 3; M365cdeGraphModule}}
        default {M365cdeAAASetup}
    }
}


Export-ModuleMember -Function M365cdeAAASetup,M365cdeGraphModule,M365cdeAzModule