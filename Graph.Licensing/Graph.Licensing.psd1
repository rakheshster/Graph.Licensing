#
# Module manifest for module 'Graph.Licensing'
#
# Generated by: Rakhesh Sasidharan
#
# Generated on: 09/20/2024
#

@{

    # Script module or binary module file associated with this manifest.
    RootModule = 'Graph.Licensing.psm1'
    
    # Version number of this module.
    ModuleVersion = '0.0.4'
    
    # Supported PSEditions
    # CompatiblePSEditions = @()
    
    # ID used to uniquely identify this module
    GUID = 'b893eb81-a255-44bc-84d6-a4aa3a0d3750'
    
    # Author of this module
    Author = 'Rakhesh Sasidharan'
    
    # Company or vendor of this module
    CompanyName = 'rakhesh.com'
    
    # Copyright statement for this module
    Copyright = '(c) Rakhesh Sasidharan. All rights reserved.'
    
    # Description of the functionality provided by this module
    Description = 'Manipulating license assignments via Graph API. Born out of frustration at having to use the M365 Admin Center to manage licenses.'
    
    # Minimum version of the PowerShell engine required by this module
    # PowerShellVersion = ''
    
    # Name of the PowerShell host required by this module
    # PowerShellHostName = ''
    
    # Minimum version of the PowerShell host required by this module
    # PowerShellHostVersion = ''
    
    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''
    
    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # ClrVersion = ''
    
    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''
    
    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.PowerShell.ConsoleGuiTools"
    )
    
    # Assemblies that must be loaded prior to importing this module
    # RequiredAssemblies = @()
    
    # Script files (.ps1) that are run in the caller's environment prior to importing this module.
    # ScriptsToProcess = @()
    
    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()
    
    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()
    
    # Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
    # NestedModules = @()
    
    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport = @(
        "Update-MgLicensingData",
        "Get-MgAssignedLicenses",
        "Update-MgAssignedLicensePlans",
        "Add-MgAssignedLicense",
        "Get-MgAvailableLicenses",
        "Remove-MgAssignedLicense"
    )
    
    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport = @()
    
    # Variables to export from this module
    #VariablesToExport = '*'
    VariablesToExport = @()
    
    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    # AliasesToExport = @()
    
    # DSC resources to export from this module
    # DscResourcesToExport = @()
    
    # List of all modules packaged with this module
    # ModuleList = @()
    
    # List of all files packaged with this module
    FileList = @('Graph.Licensing.psd1')
    
    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData = @{
    
        PSData = @{
    
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @("Graph","License","EntraID","SKU","Plan")
    
            # A URL to the license for this module.
            # LicenseUri = ''
    
            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/rakheshster/PowerShell-GraphLicensing'
    
            # A URL to an icon representing this module.
            # IconUri = ''
    
            # ReleaseNotes of this module
            ReleaseNotes = 'Fixes to Get-MgAssignedLicenses as it wasnt working as expected in case of assignments via group & direct.'
    
            # Prerelease string of this module
            # Prerelease = ''
    
            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            # RequireLicenseAcceptance = $false
    
            # External dependent modules of this module
            # ExternalModuleDependencies = @()
    
        } # End of PSData hashtable
    
    } # End of PrivateData hashtable
    
    # HelpInfo URI of this module
    # HelpInfoURI = ''
    
    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    # DefaultCommandPrefix = ''   
}

    