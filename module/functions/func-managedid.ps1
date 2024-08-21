function M365cdeMIDgraph(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
s Select Automation Account
m Manually set Managed Identity Object ID

---------------- Graph Permissions ---------------
1 Add Scope 'User.Read.All'
2 Add Scope 'User.ReadWrite.All'
3 Add Scope 'Group.Read.All'
4 Add Scope 'Group.ReadWrite.All'
5 Add Scope 'UserAuthenticationMethod.Read.All'
6 Add Scope 'AuditLog.Read.All'
7 Add Scope 'Policy.ReadWrite.ConditionalAccess'
8 Add Scope 'Mail.Send'

c Custom scope

------------ Application Permissions -------------
exo Exchange Online configurations
spo Sharepoint Online configurations

--------------- Remove Permissions ---------------
r Remove Permissions

--------------------------------------------------
b ...back to main menu

Select"){
        s {az_automation_set -breadcrumb M365cdeMIDgraph}
        m {managedid_define}
        1 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'User.Read.All'}
        2 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'User.ReadWrite.All'}
        3 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'Group.Read.All'}
        4 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'Group.ReadWrite.All'}
        5 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'UserAuthenticationMethod.Read.All'}
        6 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'AuditLog.Read.All'}
        7 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'Policy.ReadWrite.ConditionalAccess'}
        8 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope 'Mail.Send'}
        c {managedid_custom}
        exo {M365cdeMIDexo}
        spo {M365cdeMIDspo}
        r {managedid_remove}
        b {M365cdeAAASetup}
        default {M365cdeMIDgraph}
    }
}

function M365cdeMIDexo(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------- Exchange Permissions -------------

1 Add Scope 'Exchange.ManageAsApp'
2 Add Scope 'MailboxSettings.Read'
3 Add Scope 'MailboxSettings.ReadWrite'
4 Add Scope 'Mail.Read'
5 Add Scope 'Mail.ReadWrite'
6 Add Scope 'Mail.Send'
7 Add Scope 'Calendars.Read'
8 Add Scope 'Calendars.ReadWrite'
9 Add Scope 'Contacts.Read'
10 Add Scope 'Contacts.ReadWrite'

c Add custom Exchange Scope

------------------ Exchange Roles ---------------
r1 Add Role 'Exchange Recipient Administrator'
r2 Add Role 'Exchange Administrator'
rc Add custom Exchange Role

--------------------------------------------------
b ...back to previous menu

Select"){
        1 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Exchange.ManageAsApp'}
        2 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'MailboxSettings.Read'}
        3 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'MailboxSettings.ReadWrite'}
        4 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Mail.Read'}
        5 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Mail.ReadWrite'}
        6 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Mail.Send'}
        7 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Calendars.Read'}
        8 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Calendars.ReadWrite'}
        9 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Contacts.Read'}
        10 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Contacts.ReadWrite'}
        c  {managedid_customexchangescope}
        r1 {funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'Exchange Recipient Administrator'}
        r2 {funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'Exchange Administrator'}
        rc {managedid_customexchangerole}
        b {M365cdeMIDgraph}
        default {M365cdeMIDexo}
    }
}

function M365cdeMIDspo(){
    Clear-Host
    switch(Read-Host "Please select an option `
------------- Sharepoint Permissions -------------
1 Add Scope 'Sites.FullControl.All'
2 Add Scope 'Sites.Read.All'
3 Add Scope 'Sites.ReadWrite.All'
c Add custom Sharepoint Scope

---------------- Sharepoint Roles ----------------
r1 Add Role 'SharePoint Administrator'
rc Add custom Sharepoint Role

--------------------------------------------------
b ...back to previous menu

Select"){
        1 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'spo' -Scope 'Sites.FullControl.All'}
        2 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'spo' -Scope 'Sites.Read.All'}
        3 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'spo' -Scope 'Sites.ReadWrite.All'}
        c  {managedid_customsharepointscope}
        r1 {funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'SharePoint Administrator'}
        rc {managedid_customsharepointrole}
        b {M365cdeMIDgraph}
        default {M365cdeMIDspo}
    }
}


#Function for Scope assignment
function funcScopeAssignment() {
    param (
        $ManagedIdentityID,
        $Scope,
        $target
    )

    $appIds = @{
        "graph" = '00000003-0000-0000-c000-000000000000'
        "exo"   = '00000002-0000-0ff1-ce00-000000000000'
        "spo"   = '00000003-0000-0ff1-ce00-000000000000'
    }

    $appId = $appIds[$target]

    If($AutomationAccountMId) {
        $AppGraph = Get-MgServicePrincipal -Filter "AppId eq '$appId'"
        $AppRole = $AppGraph.AppRoles | Where-Object {$_.Value -eq $Scope}

        if ($appRole) {
                $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                if ($existingAppRole) { Write-Warning "The scope '$Scope' is already assigned" }
                else{
                    New-MgServicePrincipalAppRoleAssignment -PrincipalId $ManagedIdentityID -ServicePrincipalId $ManagedIdentityID -ResourceId $AppGraph.Id -AppRoleId $AppRole.Id > $null
                    $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                        if ($existingAppRole) { Write-Output "The scope '$Scope' has been assigned" }
                        else { Write-Warning "The scope '$Scope' could not be assigned" }
                }
            }
        else { Write-Warning "No App Role found for scope '$Scope'"}
    }
    else {
        Write-Warning "Managed Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
    Start-Sleep -Seconds 3
    if ($target -eq "graph") {M365cdeMIDgraph}
    elseif ($target -eq "exo") {M365cdeMIDexo}
    elseif ($target -eq "spo") {M365cdeMIDspo}
}

#Function for Managed Identity Object ID definition
function managedid_define(){
            $AutomationAccountMId = Read-Host "Enter the Object ID of your Managed Identity"
            Write-Output "Managed Identity ID is set to: $AutomationAccountMId"
            (Read-Host '
Press Enter to continue…')
            M365cdeMIDgraph
}

#Function for Custom Scope assignment
function managedid_custom(){
        $ScopeCustom = Read-Host "Enter the Scope-Name (e.g. 'User.Read.All')"
        funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'graph' -Scope $ScopeCustom
            (Read-Host '
Press Enter to continue…')
            M365cdeMIDgraph
}

#Function for Role assignment
function funcEntraRoleAssignment() {
    param (
        $ManagedIdentityID,
        $EntraRole
    )

    If($AutomationAccountMId) {

        $EntraRoleID = (Get-MgRoleManagementDirectoryRoleDefinition -Filter "DisplayName eq '$EntraRole'").Id

        if ($EntraRoleID) {
                $existingEntraRole = Get-MgRoleManagementDirectoryRoleAssignment -Filter "(PrincipalID eq '$ManagedIdentityID') and (RoleDefinitionID eq '$EntraRoleID')"
                if ($existingEntraRole) { Write-Warning "The role '$EntraRole' is already assigned" }
                else{New-MgRoleManagementDirectoryRoleAssignment -PrincipalId $ManagedIdentityID -RoleDefinitionId $EntraRoleID -DirectoryScopeId "/" ; Write-Output "The role '$EntraRole' has been assigned" }
            }
        else { Write-Warning "No Entra Role found for '$ExchangeScope'"}
    }
    else {
        Write-Warning "Managed Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
    Start-Sleep -Seconds 3
    M365cdeMIDexo
}


#Function for Custom Scope assignment EXO
function managedid_customexchangescope(){
    $ExchangeScopeCustom = Read-Host "Enter the Scope-Name (e.g. 'MailboxSettings.ReadWrite')"
    funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope $ExchangeScopeCustom
    Start-Sleep -Seconds 3
    M365cdeMIDexo
}

#Function for Custom Role assignment EXO
function managedid_customexchangerole(){
    $ExchangeRoleCustom = Read-Host "Enter the Role-Name (e.g. 'Exchange Recipient Administrator')"
    funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole $ExchangeRoleCustom
    Start-Sleep -Seconds 3
    M365cdeMIDexo
}

#Function for Custom Role assignment SPO
function managedid_customsharepointscope(){
    $SharepointScopeCustom = Read-Host "Enter the Scope-Name (e.g. 'Sites.Manage.All')"
    funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'spo' -Scope $SharepointScopeCustom
    Start-Sleep -Seconds 3
    M365cdeMIDspo
}

#Function for Custom Role assignment SPO
function managedid_customsharepointrole(){
    $SharepointRoleCustom = Read-Host "Enter the Role-Name (e.g. 'SharePoint Embedded Administrator')"
    funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole $SharepointRoleCustom
    Start-Sleep -Seconds 3
    M365cdeMIDspo
}

function managedid_remove () {
    If($AutomationAccountMId) {
        Clear-Host
        $AssignedPermissions = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationAccountMId

        # Check if there are any permissions assigned, if not, return to the main menu
        if ($AssignedPermissions.Count -eq 0) {
            Write-Output "No permissions assigned"
            Start-Sleep -Seconds 2
            M365cdeMIDgraph
        }
        
        # Add the AppDisplayName and PermissionName properties to the $AssignedPermissions object
        foreach ($permission in $AssignedPermissions) {
            $AppId = $permission.ResourceId
            $AllRoles = Get-MgServicePrincipal -Filter "Id eq '$AppId'"
            $AppDisplayname = $AllRoles.DisplayName
            $details = $AllRoles.AppRoles | Where-Object { $_.Id -eq $permission.AppRoleId }
            $permission | Add-Member -MemberType NoteProperty -Name AppDisplayName -Value $AppDisplayname
            $permission | Add-Member -MemberType NoteProperty -Name PermissionName -Value $details.Value
        }

        # List all assigned permissions
        Write-Output "Assigned permissions:"
        for ($i = 0; $i -lt $AssignedPermissions.Count; $i++) {
            $AssignedPermission = @($AssignedPermissions)[$i]
            Write-Output "$($i + 1) $($AssignedPermission.AppDisplayName) | $($AssignedPermission.PermissionName)"
        }

        # Ask the user to select a permission, or select 'a' to abort, or select 'all' permissions
        $choice = Read-Host "`nSelect an option (a to abort) - type 'all' for all permissions"
        if ($choice -match '^\d+$') { $choice = [int]$choice } # Explicitly cast to int

        # If the user selects 'a', abort the function
        if ($choice -eq 'a') { M365cdeMIDgraph }

        # If the user selects 'all', remove all permissions
        elseif ($choice -eq 'all') {
            # Ask the user to confirm the removal of all permissions by typing 'yes'
            $confirm = Read-Host "Are you sure you want to remove all permissions? Type 'yes' to confirm"
            if ($confirm -eq 'yes') {
                foreach ($permission in $AssignedPermissions) {
                    Remove-MgServicePrincipalAppRoleAssignment -AppRoleAssignmentId $permission.Id -ServicePrincipalId $AutomationAccountMId
                    Write-Output "Permission $($permission.AppDisplayName) | $($permission.PermissionName) has been removed"
                }
                Write-Output "All permissions have been removed"
            }
            else { Write-Output "No permissions have been removed" ; }
            Start-Sleep -Seconds 2
            M365cdeMIDgraph
        }
        # Elseif the user selects a number, perform the selected action (update, upgrade or remove) on the selected module
        elseif ($choice -ge 1 -and $choice -le $AssignedPermissions.Count) {
            $selectedPermission = @($AssignedPermissions)[$choice - 1]
            # Ask the user to confirm the removal of the selected permission by typing 'yes'
            $confirm = Read-Host "Are you sure you want to remove $($selectedPermission.AppDisplayName) | $($selectedPermission.PermissionName)? Type 'yes' to confirm"
            if ($confirm -eq 'yes') {
                Remove-MgServicePrincipalAppRoleAssignment -AppRoleAssignmentId $selectedPermission.Id -ServicePrincipalId $AutomationAccountMId
                Write-Output "Permission $($selectedPermission.AppDisplayName) | $($selectedPermission.PermissionName) has been removed"
            } else { Write-Output "No permissions have been removed" ; Start-Sleep -Seconds 2 ; M365cdeMIDgraph }
            Start-Sleep -Seconds 2
            managedid_remove
        }
        else {
            Write-Warning "Invalid choice. Please select a valid option."
            Start-Sleep -Seconds 2
            managedid_remove
        }
    } else {
        Write-Warning "Managed Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
}