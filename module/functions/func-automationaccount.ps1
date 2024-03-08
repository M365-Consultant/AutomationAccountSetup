function M365cdeAutomationAccount(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
s select Automation Account
--------------------------------------------------
1 Module Install 7.2
2 Module Status 7.2
3 Module Update 7.2
4 Module Remove 7.2
--------------------------------------------------
5 Module Status 5.1
6 Module Upgrade (5.1 > 7.2)
7 Module Remove 5.1

b ...back to main menu

Select"){
        s {az_automation_set}
        1 {M365cdeAutomationAccount}
        2 {az_automation_module_status -RunTimeVersion 7.2}
        3 {az_automation_module_change -RunTimeVersion 7.2 -operationMode update}
        4 {az_automation_module_change -RunTimeVersion 7.2 -operationMode remove}
        5 {az_automation_module_status -RunTimeVersion 5.1}
        6 {az_automation_module_change -RunTimeVersion 5.1 -operationMode upgrade}
        7 {az_automation_module_change -RunTimeVersion 5.1 -operationMode remove}

        b {M365cdeAAASetup}
        default {M365cdeAutomationAccount}
    }
}


function az_automation_set {
    param (
        [string]$breadcrumb
    )
    $AutomationAccountsAll = Get-AzAutomationAccount
    Clear-Host
    Write-Output "Please select a Automation Account:"
    for ($i = 0; $i -lt $AutomationAccountsAll.Count; $i++) {
        $AutomationAccount = $AutomationAccountsAll[$i]
        Write-Output "$($i + 1) $($AutomationAccount.AutomationAccountName) (RG: $($AutomationAccount.ResourceGroupName) | SubId: $($AutomationAccount.SubscriptionId))"
    }

    $choice = Read-Host "`nSelect an option (a to abort)"
    if ($choice -eq 'a') {
        M365cdeAutomationAccount
    }
    elseif ($choice -ge 1 -and $choice -le $AutomationAccountsAll.Count) {
        $selectedAutomationAccount = $AutomationAccountsAll[$choice - 1]
        $script:AutomationAccountName = $selectedAutomationAccount.AutomationAccountName
        $script:AutomationAccountRG = $selectedAutomationAccount.ResourceGroupName
        $script:AutomationAccountSubId = $selectedAutomationAccount.SubscriptionId
        $script:AutomationAccountMId = $selectedAutomationAccount.Identity.PrincipalId
        Write-Output "`nYour Selection: $($AutomationAccountName) (RG: $($AutomationAccountRG) | SubId: $($AutomationAccountSubId))"
        Start-Sleep -Seconds 1
        if ($breadcrumb -eq "M365cdeMIDgraph") {
            Write-Output "`nManaged Identity Object ID is set to: $($AutomationAccountMId)"
            Start-Sleep -Seconds 3
            M365cdeMIDgraph
        }
        else {
            Start-Sleep -Seconds 3
            M365cdeAutomationAccount
        }
    }
    else {
        Write-Output "Invalid choice. Please select a valid option."
        az_automation_set
    }
}

function az_automation_module_install {
    #global function for install/update
    param (
        [string]$moduleName,
        [string]$RunTimeVersion
    )

    New-AzAutomationModule -AutomationAccountName $AutomationAccountName `
        -ResourceGroup $AutomationAccountRG `
        -Name $moduleName `
        -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$moduleName" `
        -RuntimeVersion $RunTimeVersion
}

function az_automation_module_status {
    param (
        [string]$RunTimeVersion
    )
    $GraphModules = Get-AzAutomationModule -AutomationAccountName $AutomationAccountName -ResourceGroup $AutomationAccountRG -RuntimeVersion $RunTimeVersion | Where-Object {$_.Name -match "Microsoft.Graph"}
    $GraphModules | Select-Object Name,Version,ProvisioningState | Format-Table -AutoSize
    (Read-Host '
Press Enter to continue…')
        M365cdeAutomationAccount
}

function az_automation_module_change {
    param (
        [string]$operationMode,
        [string]$RunTimeVersion
    )

    $GraphModules = Get-AzAutomationModule -AutomationAccountName $AutomationAccountName -ResourceGroup $AutomationAccountRG -RuntimeVersion $RunTimeVersion | Where-Object {$_.Name -match "Microsoft.Graph"}
    Clear-Host
    Write-Output "Please select a Automation Account:"
    for ($i = 0; $i -lt $GraphModules.Count; $i++) {
        $GraphModule = $GraphModules[$i]
        Write-Output "$($i + 1) $($GraphModule.Name) (Version: $($GraphModule.Version))"
    }

    $choice = Read-Host "`nSelect an option (a to abort)"
    if ($choice -eq 'a') {
        M365cdeAutomationAccount
    }
    elseif ($choice -ge 1 -and $choice -le $GraphModules.Count) {
        $selectedGraphModule = $GraphModules[$choice - 1]
        if ($operationMode -eq "update"){
            az_automation_module_install -moduleName $selectedGraphModule.Name -RunTimeVersion $RunTimeVersion
            Write-Output "Updating the module to the newest version. This could take several minutes!"
        }
        elseif ($operationMode -eq "upgrade"){
            Remove-AzAutomationModule -Name $selectedGraphModule.Name -AutomationAccountName $AutomationAccountName -ResourceGroupName $AutomationAccountRG -RuntimeVersion $RunTimeVersion -Force
            az_automation_module_install -moduleName $selectedGraphModule.Name -RunTimeVersion 7.2
            Write-Output "Upgrading the v5.1 module to the newest v7.2 module. This could take several minutes!"
        }
        elseif ($operationMode -eq "remove"){
            Remove-AzAutomationModule -Name $selectedGraphModule.Name -AutomationAccountName $AutomationAccountName -ResourceGroupName $AutomationAccountRG -RuntimeVersion $RunTimeVersion
        }

        Start-Sleep -Seconds 2
        M365cdeAutomationAccount
    }
    else {
        Write-Output "Invalid choice. Please select a valid option."
        Start-Sleep -Seconds 2
        az_automation_module_change -operationMode $operationMode -RunTimeVersion $RunTimeVersion
    }
}