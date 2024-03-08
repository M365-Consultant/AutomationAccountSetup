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
        1 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'User.Read.All'}
        2 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'User.ReadWrite.All'}
        3 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'Group.Read.All'}
        4 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'Group.ReadWrite.All'}
        5 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'UserAuthenticationMethod.Read.All'}
        6 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'AuditLog.Read.All'}
        7 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'Policy.ReadWrite.ConditionalAccess'}
        8 {funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope 'Mail.Send'}
        c {managedid_custom}
        e {M365cdeMIDexo}
        b {M365cdeAAASetup}
        default {M365cdeMIDgraph}
    }
}

#Function for Graph Scope assignment
function funcGraphScopeAssignment() {
    param (
        $ManagedIdentityID,
        $GraphScope
    )

    If($AutomationAccountMId) {
        $AppGraph = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'"
        $AppRole = $AppGraph.AppRoles | Where-Object {$_.Value -eq $GraphScope}

        if ($appRole) {
                $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                if ($existingAppRole) { Write-Warning "The scope '$GraphScope' is already assigned" }
                else{
                    New-MgServicePrincipalAppRoleAssignment -PrincipalId $ManagedIdentityID -ServicePrincipalId $ManagedIdentityID -ResourceId $AppGraph.Id -AppRoleId $AppRole.Id > $null
                    $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                        if ($existingAppRole) { Write-Output "The scope '$GraphScope' has been assigned" }
                        else { Write-Warning "The scope '$GraphScope' could not be assigned" }
                }
            }
        else { Write-Warning "No App Role found for scope '$GraphScope'"}
    }
    else {
        Write-Warning "`nManaged Identity Object ID is not defined!`n`nDefine it via the options s (from a existing Automation Account) or m (manually) and try it again!"
        Start-Sleep -Seconds 5
        M365cdeMIDgraph
    }
    Start-Sleep -Seconds 3
    M365cdeMIDgraph
}

function managedid_define(){
            $AutomationAccountMId = Read-Host "Enter the Object ID of your Managed Identity"
            Write-Output "Managed Identity ID is set to: $AutomationAccountMId"
            (Read-Host '
Press Enter to continue…')
            M365cdeMIDgraph
}

function managedid_custom(){
        $ScopeCustom = Read-Host "Enter the Scope-Name (e.g. 'User.Read.All')"
        funcGraphScopeAssignment -ManagedIdentityID $AutomationAccountMId -GraphScope $ScopeCustom
            (Read-Host '
Press Enter to continue…')
            M365cdeMIDgraph
}


function M365cdeMIDexo(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
m Define Managed Identity Object ID
--------------------------------------------------
1 Add Scope 'Exchange.ManageAsApp'
2 Add Scope 'MailboxSettings.ReadWrite
3 Add Role 'Exchange Recipient Administrator'
4 Add Role 'Exchange Administrator'
--------------------------------------------------

b ...back to previous menu

Select"){
        m {managedid_define}
        1 {exo_scope_ManagAsApp}
        2 {exo_scope_MailboxSettingsReadWrite}
        3 {exo_role_RecipientAdmin}
        4 {exo_role_ExchangeAdmin}
        b {M365cdeMIDgraph}
        default {M365cdeMIDexo}
    }
}

function funcExchangeScopeAssignment() {
    param (
        $ManagedIdentityID,
        $ExchangeScope
    )

    If($AutomationAccountMId) {
        $AppGraph = Get-MgServicePrincipal -Filter "AppId eq '00000002-0000-0ff1-ce00-000000000000'"
        $AppRole = $AppGraph.AppRoles | Where-Object {$_.Value -eq $ExchangeScope}

        if ($appRole) {
                $existingAppRole = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ManagedIdentityID | Where-Object { $_.ResourceId -eq $AppGraph.Id -and $_.AppRoleId -eq $AppRole.Id }
                if ($existingAppRole) { Write-Warning "The scope '$ExchangeScope' is already assigned" }
                else{New-MgServicePrincipalAppRoleAssignment -PrincipalId $ManagedIdentityID -ServicePrincipalId $ManagedIdentityID -ResourceId $AppGraph.Id -AppRoleId $AppRole.Id ; Write-Output "The scope '$ExchangeScope' has been assigned" }
            }
        else { Write-Warning "No App Role found for scope '$ExchangeScope'"}
    }
    else {
        Write-Warning "Managed Identity Object ID is not defined! Enter the Object ID of your Managed Identity and try it again!"
        $AutomationAccountMId = Read-Host "Managed Idenitity ObjectID"
        M365cdeMIDexo
    }
}

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
        Write-Warning "Managed Identity Object ID is not defined! Enter the Object ID of your Managed Identity and try it again!"
        $AutomationAccountMId = Read-Host "Managed Idenitity ObjectID"
        M365cdeMIDexo
    }
}

function exo_scope_ManagAsApp(){
    funcExchangeScopeAssignment -ManagedIdentityID $AutomationAccountMId -ExchangeScope 'Exchange.ManageAsApp'
        (Read-Host '
Press Enter to continue…')
        M365cdeMIDexo
}

function exo_scope_MailboxSettingsReadWrite(){
    funcExchangeScopeAssignment -ManagedIdentityID $AutomationAccountMId -ExchangeScope 'MailboxSettings.ReadWrite'
        (Read-Host '
Press Enter to continue…')
        M365cdeMIDexo
}

function exo_role_RecipientAdmin(){
    funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'Exchange Recipient Administrator'
        (Read-Host '
Press Enter to continue…')
M365cdeMIDexo
}

function exo_role_ExchangeAdmin(){
    funcEntraRoleAssignment -ManagedIdentityID $AutomationAccountMId -EntraRole 'Exchange Administrator'
        (Read-Host '
Press Enter to continue…')
M365cdeMIDexo
}