function M365cdeAzModule(){
    Clear-Host
    switch(Read-Host "Please select an option `
--------------------------------------------------
1 Install Microsoft Azure Az module (CurrentUser)
2 Check installed version of Microsoft Azure Az module
3 Update Microsoft Azure Az module
--------------------------------------------------
4 Connect to Microsoft Azure
5 Disconnect from Microsoft Azure
6 Show connection state
--------------------------------------------------

b ...back to main menu

Select"){
        1 {az_install}
        2 {az_version}
        3 {az_update}
        4 {az_connect}
        5 {az_disconnect}
        6 {az_connectionstate}

        b {M365cdeAAASetup}
        default {M365cdeAzModule}
    }
}

function az_install(){
    if (Get-Module -ListAvailable -Name Az | Where-Object { $_.Path -like "$env:USERPROFILE\*"}) {
        Write-Output "The module is already installed as current user."
         Get-Module -ListAvailable -Name Az | Select-Object Name,Version,Path | Where-Object { $_.Path -like "$env:USERPROFILE\*"} | Format-List
        }
    elseif (Get-Module -ListAvailable -Name Az | Where-Object { $_.Path -like "$env:ProgramFiles\*"}) {
        Write-Warning "The module is already installed in the scope AllUsers. If you want to update this module, admin privileges are required!"
            Get-Module -ListAvailable -Name Az | Select-Object Name,Version,Path | Where-Object { $_.Path -like "$env:ProgramFiles\*"} | Format-List
        }
    else { Write-Output "Starting Module installation"
        Install-Module -Name Az -Force -Scope CurrentUser
        }
    Start-Sleep -Seconds 1
    M365cdeAzModule
}

function az_version(){
    Get-Module -ListAvailable -Name Az | Select-Object Name,Version,Path | Format-List
    (Read-Host '
Press Enter to continue…')
    M365cdeAzModule
}

function az_update(){
    if (Get-Module -ListAvailable -Name Az | Where-Object { $_.Path -like "$env:ProgramFiles\*"}) {
        Write-Warning "The module is installed in the scope AllUsers. If you want to update this module, admin privileges are required!"
    }
    Update-Module -Name Az -Force -Confirm -Scope CurrentUser
    Get-Module -ListAvailable -Name Az | Select-Object Name,Version,Path | Format-List
    Start-Sleep -Seconds 1
    M365cdeAzModule
}

function az_connect(){
    If(Get-AzContext){
        Write-Output "Closing existing connection"
        Disconnect-AzAccount
        Write-Output "Starting now connection"
        Connect-AzAccount
        Get-AzContext
    }
    else {
        Connect-AzAccount
        Get-AzContext
    }

    Start-Sleep -Seconds 2
    M365cdeAzModule
}

function az_disconnect(){
    Disconnect-AzAccount
    Start-Sleep -Seconds 2
    M365cdeAzModule
}

function az_connectionstate(){
    If(Get-AzContext){
        Get-AzContext
    }
    else {
        Write-Output "There is no active connection to Azure!"
    }
    (Read-Host '
Press Enter to continue…')
    M365cdeAzModule
}