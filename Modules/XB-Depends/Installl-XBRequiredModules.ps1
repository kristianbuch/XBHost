#requires -Version 7.5
#requires -PSEdition Core

function Install-XBRequiredModules {
    [CmdletBinding()]
    param (
        [Parameter(
            Mandatory = $false, ParameterSetName = "FromObjectSave",
            HelpMessage = "Action to perform, either 'Save' or 'Install'. 'Save' will save the module to the specified path.")]
        [Parameter(
            Mandatory = $false, ParameterSetName = "FromFileSave",
            HelpMessage = "Action to perform, either 'Save' or 'Install'. 'Save' will save the module to the specified path.")]
        [Parameter(
            Mandatory = $false, ParameterSetName = "FromObjectInstall",
            HelpMessage = "Action to perform, either 'Save' or 'Install'. 'Install' will install the module to the system.")]
        [Parameter(
            Mandatory = $false, ParameterSetName = "FromFileInstall",
            HelpMessage = "Action to perform, either 'Save' or 'Install'. 'Install' will install the module to the system.")]
        [ValidateSet("Install", "Save", IgnoreCase = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Action,

        [Parameter(
            Position = 1, Mandatory = $false,
            ValueFromPipeline = $true, ParameterSetName = "FromObjectSave",
            HelpMessage = "List of required modules to install. Each module should be a hashtable with keys: Name, Repository, AllowPreRelease, AcceptLicense, Confirm, Scope, SkipPublisherCheck, Force.")]
        [Parameter(
            Position = 1, Mandatory = $false,
            ValueFromPipeline = $true, ParameterSetName = "FromObjectInstall",
            HelpMessage = "List of required modules to install. Each module should be a hashtable with keys: Name, Repository, AllowPreRelease, AcceptLicense, Confirm, Scope, SkipPublisherCheck, Force.")]
        [array]$RequiredModules,

        [Parameter(
            Position = 1, Mandatory = $false,
            ValueFromPipeline = $true, ParameterSetName = "FromFileSave",
            HelpMessage = "Path to RequiredModules manifest file (.json or .psd1). This file should contain an array of hashtables with required keys.")]
        [Parameter(
            Position = 1, Mandatory = $false,
            ValueFromPipeline = $true, ParameterSetName = "FromFileInstall",
            HelpMessage = "Path to RequiredModules manifest file (.json or .psd1). This file should contain an array of hashtables with required keys.")]
        [string]$PathToManifest,

        [Parameter(
            Position = 2, Mandatory = $false, ParameterSetName = "FromObjectSave",
            HelpMessage = "Path to save the modules if Action is 'Save'. Default is `$env:TEMP\\Modules\\<ModuleName>.")]
        [Parameter(
            Position = 2, Mandatory = $false, ParameterSetName = "FromFileSave",
            HelpMessage = "Path to save the modules if Action is 'Save'. Default is `$env:TEMP\\Modules\\<ModuleName>.")]
        [string]$PathSave = (Join-Path -Path $env:TEMP -ChildPath "Modules")
    )

<#
    .SYNOPSIS
    Installs or saves required PowerShell modules from a predefined list or manifest file.

    .DESCRIPTION
    The Install-XBRequiredModules function processes a list of module definitions — either supplied directly as hashtables or loaded from a manifest file (.json or .psd1) — and installs or saves them depending on the selected action. It supports installation to the current user or all users, with automatic elevation handling if required.

    Modules already present in the system or in the save path will be skipped unless forced. Logging is handled via structured $PSCmdlet.Write* calls, and the function is designed for professional, script-based automation in secure and controlled environments.

    .PARAMETER RequiredModules
    A hashtable array defining the modules to install or save. This parameter is required if not using -PathToManifest.
    Each module entry must contain at least the following keys:

    - Name (string) — The module name
    - Repository (string) — The source repository (e.g. PSGallery)

    Optional keys per module:
    - AllowPreRelease (bool)
    - AcceptLicense (bool)
    - Confirm (bool)
    - Force (bool)
    - SkipPublisherCheck (bool)
    - Path (string, used for Save)
    - Scope (string, used for Install: CurrentUser or AllUsers)

    .PARAMETER PathToManifest
    Path to a .json or .psd1 file containing a module list. Must return an array of hashtables with the same structure as RequiredModules. This parameter is required if RequiredModules is not specified.

    .PARAMETER Action
    Specifies what action to perform on the module list:
    - "Install" installs modules using Install-Module.
    - "Save" saves modules using Save-Module to the path defined via -PathSave or via module.Path.

    .PARAMETER PathSave
    Only applies if Action is "Save".
    Specifies the root path to which modules are saved. Defaults to $env:TEMP\Modules\<ModuleName> if not defined.

    .INPUTS
    System.Array. Accepts hashtable arrays for -RequiredModules via pipeline or variable assignment.

    .OUTPUTS
    None. The function emits informational, debug, and error messages using $PSCmdlet streams.

    .EXAMPLE
    # Example 1: Install modules from inline hashtables
    $mods = @(
        @{ Name = "PSScriptAnalyzer"; Repository = "PSGallery"; AcceptLicense = $true },
        @{ Name = "PSReadLine"; Repository = "PSGallery" }
    )
    Install-XBRequiredModules -RequiredModules $mods -Action Install

    .EXAMPLE
    # Example 2: Save modules to local temp folder
    Install-XBRequiredModules -RequiredModules $mods -Action Save -PathSave "C:\OfflineModules"

    .EXAMPLE
    # Example 3: Install modules from a manifest file
    Install-XBRequiredModules -PathToManifest ".\RequiredModules.psd1" -Action Install

    .EXAMPLE
    # Example 4: Save prerelease module from JSON manifest
    Install-XBRequiredModules -PathToManifest ".\dev_modules.json" -Action Save -PathSave "C:\DevModules"

    .NOTES
    Name     : Install-XBRequiredModules
    Author   : Kristian Holm Buch
    Version  : 1.0
    Date     : 2025-05-17
    Project  : XBHost
    License  : CC BY-NC-ND 4.0 (https://creativecommons.org/licenses/by-nc-nd/4.0/)
    Contact  : https://linkedin.com/in/kristianbuch
    GitHub   : https://github.com/kristianbuch
    Location : Copenhagen, Denmark
    Usage    : Designed for automation, deployment, and secure module provisioning.
    Platform : Compatible with Windows PowerShell 5.1+ and PowerShell 7+

    .LINK
    https://github.com/kristianbuch/XBHost
#>
    begin {
        if ($RequiredModules) {
            $PSCmdlet.WriteDebug("Starting module installation for $($RequiredModules.Count) entries")
        } elseif ($PathToManifest) {
            $PSCmdlet.WriteDebug("Starting module installation from manifest file: $PathToManifest")
            if ($PathToManifest.EndsWith('.json')) {
                $RequiredModules = Get-Content -Path $PathToManifest | ConvertFrom-Json
            } elseif ($PathToManifest.EndsWith('.psd1')) {
                $imported = Import-PowerShellDataFile -Path $PathToManifest -SkipLimitCheck
                $RequiredModules = $imported.Modules
                Write-Host $RequiredModules
            } else {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    ([System.ArgumentException]::new("Unsupported file format. Use .json or .psd1.")),
                    "InvalidFileFormat",
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $PathToManifest
                ))
                return
            }
        } else {
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                ([System.ArgumentException]::new("Either RequiredModules or PathToManifest must be specified.")),
                "MissingRequiredParameters",
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $null
            ))
            return
        }

        switch ($Action) {
            'Save'    { $PSCmdlet.WriteDebug("Function: Save-Module (Saving modules to specified path.)") }
            'Install' { $PSCmdlet.WriteDebug("Function: Install-Module (Installing modules to system.)") }
            default {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    ([System.ArgumentException]::new("Invalid action specified. Use 'Save' or 'Install'.")),
                    "InvalidAction",
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $Action
                ))
                return
            }
        }
    }

    process {

        foreach ($Module in $RequiredModules) {
            $ModuleName = $Module.Name
            $ModuleRepository = $Module.Repository

            if ($Action -eq 'Save') {
                $moduleDir = Join-Path -Path $PathSave -ChildPath $ModuleName
                if (Test-Path $moduleDir ) {
                    $PSCmdlet.WriteDebug("Module already saved at: $moduleDir")
                    continue
                }
            }
            try {
                if (-not $Module.Name -or -not $Module.Repository) {
                    $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        ([System.ArgumentException]::new("Each module must define both 'Name' and 'Repository'.")),
                        "MissingModuleKeys",
                        [System.Management.Automation.ErrorCategory]::InvalidArgument,
                        $Module
                    ))
                }

                if ($Action -eq 'Install') {
                    $installed = Get-Module -ListAvailable -Name $Module.Name -ErrorAction SilentlyContinue
                    if ($installed) {
                        $PSCmdlet.WriteDebug("Module already installed: $($Module.Name)")
                        $PSCmdlet.WriteDebug("Module version: $($installed.Version)")
                        return
                    }
                }

                $PSCmdlet.WriteDebug("Processing module: $($Module.Name)")

                $ModuleName = $Module.Name                      ; $PSCmdlet.WriteDebug("  ModuleName      : $ModuleName")
                $ModuleRepository = $Module.Repository          ; $PSCmdlet.WriteDebug("  ModuleRepository: $ModuleRepository")

                if ($Action -eq 'Save') {
                    $AllowPreRelease = $Module.AllowPreRelease  ?? $false
                    $AcceptLicense   = $Module.AcceptLicense    ?? $false
                    $Confirm         = $Module.Confirm          ?? $false
                    $Force           = $Module.Force            ?? $true
                    $EffictivePath   = $PathSave                ?? (Join-Path -Path $env:TEMP -ChildPath "Modules\$($Module.Name)")

                    $PSCmdlet.WriteDebug("  AllowPreRelease   : $AllowPreRelease")
                    $PSCmdlet.WriteDebug("  AcceptLicense     : $AcceptLicense")
                    $PSCmdlet.WriteDebug("  Confirm           : $Confirm")
                    $PSCmdlet.WriteDebug("  Force             : $Force")
                    $PSCmdlet.WriteDebug("  Path              : $EffictivePath")
                }

                if ($Action -eq 'Install') {
                    $AllowPreRelease    = $Module.AllowPreRelease    ?? $false
                    $AcceptLicense      = $Module.AcceptLicense      ?? $false
                    $Confirm            = $Module.Confirm            ?? $false
                    $SkipPublisherCheck = $Module.SkipPublisherCheck ?? $false
                    $Scope              = $Module.Scope              ?? "CurrentUser"
                    $Force              = $Module.Force              ?? $true

                    $PSCmdlet.WriteDebug("  AllowPreRelease   : $AllowPreRelease")
                    $PSCmdlet.WriteDebug("  AcceptLicense     : $AcceptLicense")
                    $PSCmdlet.WriteDebug("  Confirm           : $Confirm")
                    $PSCmdlet.WriteDebug("  SkipPublisherCheck: $SkipPublisherCheck")
                    $PSCmdlet.WriteDebug("  Scope             : $Scope")
                    $PSCmdlet.WriteDebug("  Force             : $Force")
                }

                $PSCmdlet.WriteDebug("Installing module: $($ModuleName) from repository: $($ModuleRepository)")

                if ($Action -eq 'Install' -and $Scope -eq 'AllUsers') {
                    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
                    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
                    $IsElevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                    $IsAdminGroupMember = $identity.Groups -contains ([Security.Principal.SecurityIdentifier]::new("S-1-5-32-544"))

                    if ($IsElevated) {
                        $PSCmdlet.WriteDebug("Session is elevated. Proceeding with AllUsers install.")
                    }
                    elseif ($IsAdminGroupMember) {
                        $PSCmdlet.WriteWarning("Session is not elevated. Attempting to elevate with RunAs...")

                        $exe = (Get-Process -Id $PID).Path
                        $args = "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`""

                        try {
                            Start-Process -FilePath $exe -ArgumentList $args -Verb RunAs -WindowStyle Hidden
                            $PSCmdlet.WriteDebug("Elevation successful. Restarting script with elevated privileges.")
                            $PSCmdlet.WriteProgress(
                                [System.Management.Automation.ProgressRecord]::new(
                                    0,
                                    "Elevation in progress",
                                    "Elevating to administrator privileges. Restarting script."
                                )
                            )
                        } catch {
                            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                                $_.Exception,
                                "ElevationFailed",
                                [System.Management.Automation.ErrorCategory]::SecurityError,
                                $null
                            ))
                        }

                        return
                    }
                    else {
                        $PSCmdlet.WriteWarning("Administrator privileges are required to install modules for all users.")
                        $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                            ([System.UnauthorizedAccessException]::new("User is not a member of Administrators group. Cannot elevate.")),
                            "AdminPrivilegesRequired",
                            [System.Management.Automation.ErrorCategory]::PermissionDenied,
                            $null
                        ))
                        return
                    }
                }

                if ($Action -eq 'Install') {
                    Install-Module -Name $ModuleName `
                        -Repository $ModuleRepository `
                        -AllowPrerelease:$AllowPreRelease `
                        -AcceptLicense:$AcceptLicense `
                        -Confirm:$Confirm `
                        -SkipPublisherCheck:$SkipPublisherCheck `
                        -Scope $Scope `
                        -Force:$Force

                } elseif ($Action -eq 'Save') {
                    if (-not (Test-Path $EffictivePath)) {
                        $PSCmdlet.WriteWarning("Path does not exist. Creating directory: $EffictivePath")
                        New-Item -Path $EffictivePath -ItemType Directory -Force | Out-Null
                        $PSCmdlet.WriteDebug("Directory created: $EffictivePath")

                        if (-not (Test-Path $EffictivePath)) {
                            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                                ([System.IO.DirectoryNotFoundException]::new("Failed to create directory: $EffictivePath")),
                                "DirectoryCreationFailed",
                                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                                $EffictivePath
                            ))
                            return
                        }
                    }

                    Save-Module -Name $ModuleName `
                        -Repository $ModuleRepository `
                        -AllowPrerelease:$AllowPreRelease `
                        -AcceptLicense:$AcceptLicense `
                        -Confirm:$Confirm `
                        -Force:$Force `
                        -Path $EffictivePath

                }

                $PSCmdlet.WriteDebug("Module $ModuleName installed/saved successfully.")

            } catch [System.ArgumentException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::InvalidArgument,
                    $ModuleName
                ))
            } catch [System.UnauthorizedAccessException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::PermissionDenied,
                    $ModuleName
                ))
            } catch [System.IO.IOException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::WriteError,
                    $ModuleName
                ))
            } catch [System.Security.SecurityException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::SecurityError,
                    $ModuleName
                ))
            } catch [System.Management.Automation.RuntimeException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::InvalidOperation,
                    $ModuleName
                ))
            } catch [System.Net.WebException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::ConnectionError,
                    $ModuleName
                ))
            } catch [System.InvalidOperationException] {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::InvalidOperation,
                    $ModuleName
                ))
            } catch {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    $_.Exception,
                    "ModuleInstallationFailed",
                    [System.Management.Automation.ErrorCategory]::NotSpecified,
                    $ModuleName
                ))
            }
        }
    }

    end {
        $PSCmdlet.WriteDebug("Finished processing all required modules")
    }
}
