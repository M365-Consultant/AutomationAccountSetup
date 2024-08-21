function M365cdeMaester(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
s select Automation Account

-------------- Runtime version 7.2 ---------------
1 Module Install Pester
2 Module Install Maester
3 Module Install Microsoft.Graph.Authentication
4 Module Status
5 Module Update
6 Module Remove

------------------ Permissions -------------------
7 Permissions Check
8 Permissions Assignment

--------------------------------------------------
b ...back to main menu

Select"){
        s {az_automation_set -breadcrumb "M365cdeMaester"}
        1 {az_automation_module_install -moduleName "Pester" - -RunTimeVersion 7.2 -breadcrumb "M365cdeMaester"}
        2 {az_automation_module_install -moduleName "Maester" - -RunTimeVersion 7.2 -breadcrumb "M365cdeMaester"}
        3 {az_automation_module_install -moduleName "Microsoft.Graph.Authentication" -RunTimeVersion 7.2 -breadcrumb "M365cdeMaester"}
        4 {az_automation_module_status -RunTimeVersion 7.2 -filter "Pester|Maester|Microsoft.Graph.Authentication" -breadcrumb "M365cdeMaester"}
        5 {az_automation_module_change -RunTimeVersion 7.2 -Mode update -filter "Pester|Maester|Microsoft.Graph.Authentication" -breadcrumb "M365cdeMaester"}
        6 {az_automation_module_change -RunTimeVersion 7.2 -Mode remove -filter "Pester|Maester|Microsoft.Graph.Authentication" -breadcrumb "M365cdeMaester"}
        7 {managedid_Maester -ManagedIdentityID $AutomationAccountMId -mode check}
        8 {managedid_Maester -ManagedIdentityID $AutomationAccountMId -mode assign}


        b {M365cdeAAASetup}
        default {M365cdeAutomationAccount}
    }
}

function managedid_Maester(){
    param (
        $ManagedIdentityID,
        $mode
    )
    
    Clear-Host

    # Required Permissions for Maester
    $Permissions = @("Directory.Read.All","Policy.Read.All","Reports.Read.All","DirectoryRecommendations.Read.All","PrivilegedAccess.Read.AzureAD","IdentityRiskEvent.Read.All","RoleEligibilitySchedule.Read.Directory","Policy.Read.ConditionalAccess","Mail.Send")

    $appId = '00000003-0000-0000-c000-000000000000'

    If($AutomationAccountMId) {
        # Get the Service Principal for the Managed Identity
        $AppGraph = Get-MgServicePrincipal -Filter "AppId eq '$appId'"

        # check for each scope if it is already assigned and output the result as a table
        $PermissionsCheck = @()
        foreach ($item in $Permissions) {
            if ($mode -eq "check") { Write-Output "Checking permissions for scope '$item':" }
            $AppRole = $AppGraph.AppRoles | Where-Object {$_.Value -eq $item}
            if ($appRole) {
                $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                if ($existingAppRole) { $PermissionsCheck += [PSCustomObject]@{Scope=$item;ApproleID=$AppRole.Id;Assigned="Yes"} }
                else {
                    $PermissionsCheck += [PSCustomObject]@{Scope=$item;ApproleID=$AppRole.Id;Assigned="No"}
                    if ($mode -eq "assign") {
                        Write-Output "Assigning App Role for scope '$item'…"
                        New-MgServicePrincipalAppRoleAssignment -PrincipalId $ManagedIdentityID -ServicePrincipalId $ManagedIdentityID -ResourceId $AppGraph.Id -AppRoleId $AppRole.Id > $null
                        $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                        if ($existingAppRole) { Write-Output "The scope '$item' has been assigned" }
                        else { Write-Warning "The scope '$item' could not be assigned" }
                    }
                }
            }
            else { Write-Warning "No App Role found for scope '$item'"}
        }
        
        if ($mode -eq "check") {
            Clear-Host
            $PermissionsCheck | Format-Table -AutoSize
        } elseif ($mode -eq "assign") {
            Write-Output "`nPermissions have been assigned.`nChecking the permissions assignment.`n`nPlease wait..."
            Start-Sleep -Seconds 3
            managedid_Maester -ManagedIdentityID $AutomationAccountMId -mode check
        }

    } else {
        Write-Warning "`nManaged Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
    (Read-Host '
Press Enter to continue…')
    M365cdeMaester
}