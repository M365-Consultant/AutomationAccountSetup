#
# Module manifest for module 'M365cde.AutomationAccountSetup'
#
# Generated by: Dominik Gilgen
#
# Generated on: 2024-03-08
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'M365cde.AutomationAccountSetup.psm1'

# Version number of this module.
ModuleVersion = '0.1.0'

# Supported PSEditions
 CompatiblePSEditions = @( 'Desktop', 'Core' )

# ID used to uniquely identify this module
GUID = 'c1d0c18b-571e-4f07-a0b5-3e6a8791bc00'

# Author of this module
Author = 'Dominik Gilgen'

# Company or vendor of this module
CompanyName = 'Dominik Gilgen (Personal)'

# Copyright statement for this module
Copyright = '(c) 2024 Dominik Gilgen. All rights reserved.'

# Description of the functionality provided by this module
Description = 'This module helps you set up an Azure Automation Account.'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '5.1'

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @('M365cdeAAASetup')

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = @()

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = @('PowerShell', 'Module', 'Automation')

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/M365-Consultant/AzureAutomationSetup/blob/main/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/M365-Consultant/AzureAutomationSetup'

        # Pre-release string for the module
        Prerelease = 'alpha'

        # ReleaseNotes of this module
        ReleaseNotes = @'
v0.1.0 - 2024-08-21
Measter: Added a new feature for handling the Maester Test Framework (beta).
Automation Account Function: Changed the function to use a filter parameter for the module status and update functions. Added breadcrumb-support, required for the new Maester feature. Added support for installing a specific version of a module. Fixed the remove-function when selecting all modules. Fixed a bug within the module selection, when there was only one module installed
Managed Identity Function: Added a function to remove permissions.
Some code improvements and cosmetic changes.

v0.0.6 - 2024-05-28
Automation Account Function: Added a new feature to update all modules at once. Added a new feature to install modules from the PS Gallery. Explicitly cast the menu choice to int. Multiple code improvements.
Managed Identity Function: Fixed a wrong link within the menu.

v0.0.5 - 2024-04-17
Implemented SharePoint scope assignment.

v0.0.4 - 2024-03-19
Fixed Exchange Online Scopes.

v0.0.3 - 2024-03-19
Simplyfied the code for scope assignment.

v0.0.2 - 2024-03-18
Changed the code for scope assignment.

v0.0.1 - 2024-03-08
Initial release of the module.
'@

    } # End of PSData hashtable

} # End of PrivateData hashtable

}