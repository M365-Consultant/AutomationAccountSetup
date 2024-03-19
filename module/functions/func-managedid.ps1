function M365cdeMIDgraph(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
s Select Automation Account
m Manually set Managed Identity Object ID
--------------------------------------------------
1 Add Scope 'User.Read.All'
2 Add Scope 'User.ReadWrite.All'
3 Add Scope 'Group.Read.All'
4 Add Scope 'Group.ReadWrite.All'
5 Add Scope 'UserAuthenticationMethod.Read.All'
6 Add Scope 'AuditLog.Read.All'
7 Add Scope 'Policy.ReadWrite.ConditionalAccess'
8 Add Scope 'Mail.Send'
--------------------------------------------------
c Custom scope
--------------------------------------------------
e Exchange Online configurations
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
        e {M365cdeMIDexo}
        b {M365cdeAAASetup}
        default {M365cdeMIDgraph}
    }
}

function M365cdeMIDexo(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
1 Add Scope 'Exchange.ManageAsApp'
2 Add Scope 'MailboxSettings.Read'
3 Add Scope 'MailboxSettings.ReadWrite'
4 Add Scope 'Mail.Read'
5 Add Scope 'Mail.ReadBasic'
6 Add Scope 'Mail.ReadWrite'
7 Add Scope 'Mail.Send'
8 Add Scope 'Calenders.Read'
9 Add Scope 'Calenders.ReadWrite'
10 Add Scope 'Contacts.Read'
11 Add Scope 'Contacts.ReadWrite'
c Add custom Exchange Scope
--------------------------------------------------
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
        5 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Mail.ReadBasic'}
        6 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Mail.ReadWrite'}
        7 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Mail.Send'}
        8 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Calenders.Read'}
        9 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Calenders.ReadWrite'}
        10 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Contacts.Read'}
        11 {funcScopeAssignment -ManagedIdentityID $AutomationAccountMId -target 'exo' -Scope 'Contacts.ReadWrite'}
        c  {managedid_customexchangescope}
        r1 {funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'Exchange Recipient Administrator'}
        r2 {funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'Exchange Administrator'}
        rc {managedid_customexchangerole}
        b {M365cdeMIDgraph}
        default {M365cdeMIDexo}
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
        Write-Warning "`nManaged Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
    Start-Sleep -Seconds 3
    if ($target -eq "graph") {M365cdeMIDgraph}
    elseif ($target -eq "exo") {M365cdeMIDexo}
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
        Write-Warning "`nManaged Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
    Start-Sleep -Seconds 3
    M365cdeMIDexo
}

#Function for Custom Role assignment EXO
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