function M365cdeAutomationAccount(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
s select Automation Account
--------------------------------------------------
1 Module Install 7.2 (from PSGallery)
2 Module Status 7.2
3 Module Update 7.2
4 Module Remove 7.2
--------------------------------------------------
5 Module Install 5.1 (from PSGallery)
6 Module Status 5.1
7 Module Update 5.1
8 Module Upgrade (5.1 > 7.2)
9 Module Remove 5.1

b ...back to main menu

Select"){
        s {az_automation_set}
        1 {az_automation_module_psginstall -RunTimeVersion 7.2}
        2 {az_automation_module_status -RunTimeVersion 7.2}
        3 {az_automation_module_change -RunTimeVersion 7.2 -Mode update}
        4 {az_automation_module_change -RunTimeVersion 7.2 -Mode remove}
        5 {az_automation_module_psginstall -RunTimeVersion 5.1}
        6 {az_automation_module_status -RunTimeVersion 5.1}
        7 {az_automation_module_change -RunTimeVersion 5.1 -Mode update}
        8 {az_automation_module_change -RunTimeVersion 5.1 -Mode upgrade}
        9 {az_automation_module_change -RunTimeVersion 5.1 -Mode remove}

        b {M365cdeAAASetup}
        default {M365cdeAutomationAccount}
    }
}


function az_automation_set {
    # Function to select the Automation Account
    param (
        [string]$breadcrumb
    )
    $AutomationAccountsAll = Get-AzAutomationAccount # Get all Automation Accounts
    Clear-Host
    Write-Output "Please select a Automation Account:"
    # Create a list of all Automation Accounts with the Subscription ID and Resource Group Name
    for ($i = 0; $i -lt $AutomationAccountsAll.Count; $i++) {
        $AutomationAccount = $AutomationAccountsAll[$i]
        Write-Output "$($i + 1) $($AutomationAccount.AutomationAccountName) (RG: $($AutomationAccount.ResourceGroupName) | SubId: $($AutomationAccount.SubscriptionId))"
    }

    # Ask the user to select an Automation Account
    $choice = Read-Host "`nSelect an option (a to abort)"
    if ($choice -match '^\d+$') { $choice = [int]$choice } # Explicitly cast to int

    # If the user selects 'a', abort the function
    if ($choice -eq 'a') {
        M365cdeAutomationAccount
    }
    # Elseif the user selects a number, set the Automation Account variables and go to the next function
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
    # Else, if the user selects an invalid option, show an error message and restart the function
    else {
        Write-Output "Invalid choice. Please select a valid option."
        az_automation_set
    }
}

function az_automation_module_install {
    #Function for install/update a module from the PSGallery
    param (
        [string]$moduleName,
        [string]$RunTimeVersion,
        [switch]$NoExit = $false
    )

    # Install the module from the PSGallery
    New-AzAutomationModule -AutomationAccountName $AutomationAccountName `
        -ResourceGroup $AutomationAccountRG `
        -Name $moduleName `
        -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$moduleName" `
        -RuntimeVersion $RunTimeVersion

    # If the NoExit switch is not set, ask the user to press Enter to continue and go back to the main menu (this is required for the update all process)
    If ($NoExit -eq $false) {
    (Read-Host '
Press Enter to continue…')
         M365cdeAutomationAccount
    }
}

function az_automation_module_psginstall {
    # Function to install a module asking the user for the module name
    param (
        [string]$RunTimeVersion
    )
    $psgName = Read-Host "Please enter the PowerShell Gallery module name to install:"
    az_automation_module_install -moduleName $psgName -RunTimeVersion $RunTimeVersion
}

function az_automation_module_status {
    # Function to show the status of the Microsoft.Graph and M365cde modules
    param (
        [string]$RunTimeVersion
    )
    $GraphModules = Get-AzAutomationModule -AutomationAccountName $AutomationAccountName -ResourceGroup $AutomationAccountRG -RuntimeVersion $RunTimeVersion | Where-Object {($_.Name -match "Microsoft.Graph") -or ($_.Name -match "M365cde")}
    $GraphModules | Select-Object Name,Version,ProvisioningState | Format-Table -AutoSize
    (Read-Host '
Press Enter to continue…')
        M365cdeAutomationAccount
}

function az_automation_module_change {
    # Function to update, upgrade or remove a module
    param (
        [string]$Mode,
        [string]$RunTimeVersion
    )

    # Get all Microsoft.Graph and M365cde modules from the Automation Account
    $GraphModules = Get-AzAutomationModule -AutomationAccountName $AutomationAccountName -ResourceGroup $AutomationAccountRG -RuntimeVersion $RunTimeVersion | Where-Object {($_.Name -match "Microsoft.Graph") -or ($_.Name -match "M365cde")}
    
    # Check if there is an update available on the PSgallery and add the newest version to the object as "NewestVersion"
    Clear-Host
    Write-Output "Checking for updates. Sometimes this can take a while, please wait..."
    $GraphModulesCount = $GraphModules.Count
    foreach ($index in 0..($GraphModulesCount - 1)) {
        $GraphModule = $GraphModules[$index]
        $newestVersion = (Find-Module -Name $GraphModule.Name -Repository PSGallery).Version
        $GraphModule | Add-Member -MemberType NoteProperty -Name NewestVersion -Value $newestVersion

        # Output progress to the user
        $progress = ($index + 1)
        Write-Output "$progress of $GraphModulesCount ✅ $($GraphModule.Name)"
    }

    Clear-Host
    # Show the user the available modules and the newest version if it is higher than the installed version
    Write-Output "Please select a module:"
    for ($i = 0; $i -lt $GraphModules.Count; $i++) {
        $GraphModule = $GraphModules[$i]
        Write-Output "$($i + 1) $($GraphModule.Name) (Version: $($GraphModule.Version))$(if ($GraphModule.NewestVersion -gt $GraphModule.Version) { ' ----> Newest Version: ' + $GraphModule.NewestVersion })"
    }
    
    # Ask the user to select a module, or select 'a' to abort, or select 'all' to update all modules
    $choice = Read-Host "`nSelect an option (a to abort) - type 'all' to update all modules"
    if ($choice -match '^\d+$') { $choice = [int]$choice } # Explicitly cast to int
    
    # If the user selects 'a', abort the function
    if ($choice -eq 'a') {
        M365cdeAutomationAccount
    }
    # Elseif the user selects 'all', update all modules where the newest version is higher than the installed version. Skip the modules where the ProvisioningState is not "Succeeded"
    elseif ($choice -eq 'all') {
        foreach ($GraphModule in $GraphModules) {
            if ($GraphModule.NewestVersion -gt $GraphModule.Version -and $GraphModule.ProvisioningState -eq "Succeeded") {
                az_automation_module_install -moduleName $GraphModule.Name -RunTimeVersion $RunTimeVersion -NoExit
                Write-Output "Updating $($GraphModule.Name) to the newest version. This could take several minutes!"
            }
        }
        (Read-Host '
All Modules will be updated in the background.  press Enter to continue…')
        M365cdeAutomationAccount
    }
    # Elseif the user selects a number, perform the selected action (update, upgrade or remove) on the selected module
    elseif ($choice -ge 1 -and $choice -le $GraphModules.Count) {
        $selectedGraphModule = $GraphModules[$choice - 1]
        if ($Mode -eq "update"){
            az_automation_module_install -moduleName $selectedGraphModule.Name -RunTimeVersion $RunTimeVersion
            Write-Output "Updating the module to the newest version. This could take several minutes!"
        }
        elseif ($Mode -eq "upgrade"){
            Remove-AzAutomationModule -Name $selectedGraphModule.Name -AutomationAccountName $AutomationAccountName -ResourceGroupName $AutomationAccountRG -RuntimeVersion $RunTimeVersion -Force
            az_automation_module_install -moduleName $selectedGraphModule.Name -RunTimeVersion 7.2
            Write-Output "Upgrading the v5.1 module to the newest v7.2 module. This could take several minutes!"
        }
        elseif ($Mode -eq "remove"){
            Remove-AzAutomationModule -Name $selectedGraphModule.Name -AutomationAccountName $AutomationAccountName -ResourceGroupName $AutomationAccountRG -RuntimeVersion $RunTimeVersion
        }

        Start-Sleep -Seconds 2
        M365cdeAutomationAccount
    }
    # Else, if the user selects an invalid option, show an error message and restart the function
    else {
        Write-Output "Invalid choice. Please select a valid option."
        Start-Sleep -Seconds 2
        az_automation_module_change -Mode $Mode -RunTimeVersion $RunTimeVersion
    }
}