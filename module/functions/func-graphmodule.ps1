function M365cdeGraphModule(){
    Clear-Host
    switch(Read-Host "Please select an option `
----------------- Stable Module ------------------
1 Install Microsoft Graph module (CurrentUser)
2 Check installed version of Microsoft Graph module
3 Update Microsoft Graph module (CurrentUser)

----------------- Beta Module --------------------
1beta Install Microsoft Graph Beta module (CurrentUser)
2beta Check installed version of Microsoft Graph Beta module
3beta Update Microsoft Graph Beta module (CurrentUser)

------------------ Connection --------------------
4 Connect to Microsoft Graph
5 Disconnect from Microsoft Graph
6 Show connection state
7 Show active scopes

--------------------------------------------------
b ...back to main menu

Select"){
        1 {mggraph_install}
        2 {mggraph_version}
        3 {mggraph_update}
        1beta {mggraph_beta_install}
        2beta {mggraph_beta_version}
        3beta {mggraph_beta_update}
        4 {mggraph_connect}
        5 {mggraph_disconnect}
        6 {mggraph_connectionstate}
        7 {mggraph_activescopes}

        b {M365cdeAAASetup}
        default {M365cdeGraphModule}
    }
}

function mggraph_install(){
            if (Get-Module -ListAvailable -Name Microsoft.Graph | Where-Object { $_.Path -like "$env:USERPROFILE\*"}) {
                Write-Output "The module is already installed as current user."
                 Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object Name,Version,Path | Where-Object { $_.Path -like "$env:USERPROFILE\*"} | Format-List
                }
            elseif (Get-Module -ListAvailable -Name Microsoft.Graph | Where-Object { $_.Path -like "$env:ProgramFiles\*"}) {
                Write-Warning "The module is already installed in the scope AllUsers. If you want to update this module, admin privileges are required!"
                    Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object Name,Version,Path | Where-Object { $_.Path -like "$env:ProgramFiles\*"} | Format-List
                }
            else { Write-Output "Starting Module installation"
                Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser
                }
            Start-Sleep -Seconds 1
            M365cdeGraphModule
}

function mggraph_version(){
            Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object Name,Version,Path | Format-List
            (Read-Host '
Press Enter to continue…')
            M365cdeGraphModule
}

function mggraph_update(){
            if (Get-Module -ListAvailable -Name Microsoft.Graph | Where-Object { $_.Path -like "$env:ProgramFiles\*"}) {
                Write-Warning "The module is installed in the scope AllUsers. If you want to update this module, admin privileges are required!"
            }
            Update-Module -Name Microsoft.Graph -Force -Confirm -Scope CurrentUser
            Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object Name,Version,Path | Format-List
            Start-Sleep -Seconds 1
            M365cdeGraphModule
}

function mggraph_beta_install(){
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Where-Object { $_.Path -like "$env:USERPROFILE\*"}) {
        Write-Output "The module is already installed as current user."
         Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Select-Object Name,Version,Path | Where-Object { $_.Path -like "$env:USERPROFILE\*"} | Format-List
        }
    elseif (Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Where-Object { $_.Path -like "$env:ProgramFiles\*"}) {
        Write-Warning "The module is already installed in the scope AllUsers. If you want to update this module, admin privileges are required!"
            Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Select-Object Name,Version,Path | Where-Object { $_.Path -like "$env:ProgramFiles\*"} | Format-List
        }
    else { Write-Output "Starting Module installation"
        Install-Module -Name Microsoft.Graph.Beta -Force -Scope CurrentUser
        }
    Start-Sleep -Seconds 1
    M365cdeGraphModule
}

function mggraph_beta_version(){
    Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Select-Object Name,Version,Path | Format-List
    (Read-Host '
Press Enter to continue…')
    M365cdeGraphModule
}

function mggraph_beta_update(){
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Where-Object { $_.Path -like "$env:ProgramFiles\*"}) {
        Write-Warning "The module is installed in the scope AllUsers. If you want to update this module, admin privileges are required!"
    }
    Update-Module -Name Microsoft.Graph.Beta -Force -Confirm -Scope CurrentUser
    Get-Module -ListAvailable -Name Microsoft.Graph.Beta | Select-Object Name,Version,Path | Format-List
    Start-Sleep -Seconds 1
    M365cdeGraphModule
}

function mggraph_connect(){
            If(Get-MgContext){
                Write-Output "Closing existing connection"
                Disconnect-MgGraph
                Write-Output "Starting now connection"
                Connect-MgGraph -Scopes "Application.Read.All","AppRoleAssignment.ReadWrite.All,RoleManagement.ReadWrite.Directory,User.ReadWrite.All,MailboxSettings.ReadWrite"
                Get-MgContext
            }
            else {
                Connect-MgGraph -Scopes "Application.Read.All","AppRoleAssignment.ReadWrite.All,RoleManagement.ReadWrite.Directory,User.ReadWrite.All,MailboxSettings.ReadWrite"
                Get-MgContext
            }

            Start-Sleep -Seconds 2
            M365cdeGraphModule
}

function mggraph_disconnect(){
            Disconnect-MgGraph
            Start-Sleep -Seconds 2
            M365cdeGraphModule
}

function mggraph_connectionstate(){
            If(Get-MgContext){
                Get-MgContext
            }
            else {
                Write-Output "There is no active connection to Microsoft Graph!"
            }
            (Read-Host '
Press Enter to continue…')
            M365cdeGraphModule
}

function mggraph_activescopes(){

            If(Get-MgContext){
                $activescopes = Get-MgContext
                Write-Output "Those scopes are currently active:"
                Write-Output $activescopes.Scopes | Format-List
            }
            else {
                Write-Output "There is no active connection to Microsoft Graph!"
            }
            (Read-Host '
Press Enter to continue…')
            M365cdeGraphModule
}
