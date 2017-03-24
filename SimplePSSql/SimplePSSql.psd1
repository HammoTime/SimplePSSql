#
# Module manifest for module 'SimplePSSql'
#

@{

# Version number of this module.
ModuleVersion = '1.0.0.0'

# ID used to uniquely identify this module
GUID = '485558e3-ba94-44cb-8ca8-f4b2db0690f3'

# Author of this module
Author = 'Adam Hammond'

# Company or vendor of this module
CompanyName = 'None'

# Copyright statement for this module
Copyright = '(c) Adam Hammond. All rights reserved.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '4.0'

# List of all modules packaged with this module
ModuleList = @('SimplePSSql')

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
NestedModules = @('.\SimplePSSql.psm1', '.\SimplePSSqlUtilities.psm1')

# Functions to export from this module
FunctionsToExport = @('Select-SqlScalar', 'Select-SqlRows', 'Update-SqlTable', 'Test-SqlConnection', 'Update-SimplePSSql')

# HelpInfo URI of this module
HelpInfoURI = 'https://github.com/HammoTime/SimplePSSql'

}