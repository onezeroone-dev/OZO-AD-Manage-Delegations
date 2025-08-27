#Requires -Modules ActiveDirectory,DFSN,DFSR,DSACL,GroupPolicy,ImportExcel,OZO,OZOAD,OZOFiles,OZOLogger

<#PSScriptInfo
    .VERSION 1.0.0
    .GUID 069ad55f-163a-4900-b35b-2a1100d64e81
    .AUTHOR Andy Lievertz <alievertz@onezeroone.dev>
    .COMPANYNAME One Zero One
    .COPYRIGHT (c) 2025
    .TAGS 
    .LICENSEURI https://github.com/onezeroone-dev/OZO-AD-Manage-Delegations/blob/main/LICENSE
    .PROJECTURI https://github.com/onezeroone-dev/OZO-AD-Manage-Delegations
    .ICONURI 
    .EXTERNALMODULEDEPENDENCIES ActiveDirectory,DFSN,DFSR,DSACL,GroupPolicy,ImportExcel
    .REQUIREDSCRIPTS 
    .EXTERNALSCRIPTDEPENDENCIES 
    .RELEASENOTES https://github.com/onezeroone-dev/OZO-AD-Manage-Delegations/blob/main/CHANGELOG.md
#>

<# 
    .SYNOPSIS
    See description.
    .DESCRIPTION 
    Creates AD delegations based on a configuration file. This script can create OU Delegations, apply permissions to GPOs, grant access to DFSN roots, grant access to DFSN folders, and create delegations to DFSR replication groups.
    .PARAMETER Configuration
    Path to the JSON configuration file. Defaults to "ad-create-delegations.json" in the same directory as the script.
    .PARAMETER OutDir
    Path for the Excel report. Defaults to the current directory.
    .LINK
    https://github.com/onezeroone-dev/OZO-AD-Manage-Delegations/blob/main/README.md
    .LINK
    https://github.com/SimonWahlin/DSACL/blob/master/docs
    .LINK
    https://learn.microsoft.com/en-us/powershell/module/grouppolicy/set-gppermission?view=windowsserver2022-ps
    .LINK
    https://learn.microsoft.com/en-us/powershell/module/dfsn/grant-dfsnaccess?view=windowsserver2022-ps
    .LINK
    https://learn.microsoft.com/en-us/powershell/module/dfsr/grant-dfsrdelegation?view=windowsserver2022-ps
    .NOTES
    Run this script as a user with rights to create AD delegations (e.g., a Domain Admin) from within a writable directory.
#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess = $true)] Param (
    [Parameter(Mandatory=$false,HelpMessage="Path to the JSON configuration file")][String]$Configuration = (Join-Path -Path $PSScriptRoot -ChildPath "ad-create-delegations.json"),
    [Parameter(Mandatory=$false,HelpMessage="Path for the Excel report")][String]$OutDir = (Get-Location)
)

# CLASSES
Class Main {
    # PROPERTIES: Booleans, Hashtables, Strings
    [Boolean]   $Validates = $true
    [String]    $excelPath = $null
    [String]    $jsonPath  = $null
    [String]    $outDir    = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $Json     = $null
    [PSCustomObject] $ozoLogger = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[PSCustomObject]] $ouDelegations         = @()
    [System.Collections.Generic.List[PSCustomObject]] $gpoPermissions        = @()
    [System.Collections.Generic.List[PSCustomObject]] $dfsnRootPermissions   = @()
    [System.Collections.Generic.List[PSCustomObject]] $dfsnFolderPermissions = @()
    [System.Collections.Generic.List[PSCustomObject]] $dfsrPermissions       = @()
    # METHODS: Constructor method
    Main($Configuration,$OutDir) {
        # Set properties
        $this.jsonPath = $Configuration
        $this.outDir   = $OutDir
        # Create a ozoLogger object
        $this.ozoLogger = (New-OZOLogger)
        # Declare ourselves to the world
        $this.ozoLogger.Write("Starting process.","Information")
        # And the results of ValidateConfiguration and ValidateEnvironment to set validates
        If (($this.ValidateConfiguration() -And $this.ValidateEnvironment()) -eq $true) {
            # Configuration and environment validate; call the permissions methods
            $this.CreateOUDelegations()
            $this.SetGPOPermissions()
            $this.GrantDFSNRootPermissions()
            $this.GrantDFSNFolderPermissions()
            $this.GrantDFSRPermissions()
        } Else {
            $this.Validates = $false
        }
        # Report
        $this.Report()
        $this.ozoLogger.Write("Process complete.","Information")
    }
    # METHODS: Configuration validation method
    Hidden [Boolean] ValidateConfiguration() {
        # control variable
        [Boolean]$Return = $true
        # Check that the jsonPath is valid
        Try {
            Test-Path -Path $this.jsonPath -ErrorAction Stop
            # Success; attempt to read the JSON
            Try {
                $this.Json = Get-Content $this.jsonPath -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                # Success (able to read JSON)
                $this.ozoLogger.Write("Configuration validates.","Information")
            } Catch {
                # Failure (unable to read JSON)
                $this.ozoLogger.Write(("Invalid JSON in " + $this.jsonPath + "."),"Error")
                $Return = $false
            }
        } Catch {
            # JSON path is not valid
            $this.ozoLogger.Write(("Could not read configuration file " + $this.jsonPath + "."),"Error")
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: Environment validation method
    Hidden [Boolean] ValidateEnvironment() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if outDir is writable
        If ((Test-OZOPath -Path $this.outDir -Writable) -eq $true) {
            # outDir is writable
            $this.excelPath = (Join-Path -Path $this.outDir -ChildPath ((Get-OZO8601Date -Time) + "-ad-create-delegations-report.xlsx"))
            $this.ozoLogger.Write(("Using " + $this.excelPath + " for the Excel report"),"Information")
        } Else {
            # outDir is not writable; determine if current location is writable
            $this.ozoLogger.Write("Provided output directory is not writable.","Warning")
            If ((Test-SHPathWritable -Path (Get-Location)) -eq $true) {
                # current directory is writable
                $this.excelPath = (Join-Path -Path (Get-Location) -ChildPath ((Get-SH8601Date -Time) + "-ad-create-delegations-report.xlsx"))
                $this.ozoLogger.Write(("Using " + $this.excelPath + " for the Excel report"),"Information")
            } Else {
                # current directory is not writable
                $this.ozoLogger.Write("Current directory is not writable; cannot proceed.","Error")
                $return = $false
            }
        }
        # Return
        return $Return
    }
    # METHODS: CreateDelegations method
    Hidden [Void] CreateOUDelegations() {
        # Determine if there are OU Delegations to process
        If (($this.Json.ADOUDelegations).Count -gt 0) {
            # There are OU Delegations to process; report
            $this.ozoLogger.Write("Processing Delegations.","Information")
            # Iterate through the OU Delegations
            ForEach ($delegation in $this.Json.ADOUDelegations) {
                # Report
                $this.ozoLogger.Write(("Processing " + $delegation.Description + "."),"Information")
                # Iterate through the OUs
                ForEach ($ouDN in $delegation.OUs) {
                    # Iterate through the Identities
                    ForEach ($identity in $delegation.Identities) {
                        # Iterate through the Permissions
                        ForEach ($permission in $delegation.Permissions) {
                            # Create an OUDelegation object for this Delegation's OU + Identity + Permission
                            $this.ouDelegations.Add(([OUDelegation]::new($ouDN,$identity,$permission)))
                        }
                    }
                }
            }
        } Else {
            # There are no OU Delegations to process
            $this.ozoLogger.Write("No delegations to process.","Warning")
        }
    }
    # METHODS: SetGPOPermissions method
    Hidden [Void] SetGPOPermissions() {
        # Determine if there are GPO Permissions to process
        If (($this.Json.ADGPOPermissions).Count -gt 0) {
            # There are GPO Permissiont to process; report
            $this.ozoLogger.Write("Processing GPO Permissions.","Information")
            # Iterate through the Permissions
            ForEach ($gpoPermission in $this.Json.ADGPOPermissions) {
                # Report
                $this.ozoLogger.Write(("Processing " + $gpoPermission.Description + "."),"Information")
                # Iterate through the GPO Names
                ForEach ($gpoName in $gpoPermission.GPONames) {
                    # Iterate through the Group Names
                    ForEach ($group in $gpoPermission.GroupNames) {
                        # Iterate through the Permissions
                        ForEach ($permission in $gpoPermission.Permissions) {
                            # Create a GPOPermissions object for this GPO Name + Group Name + Permission
                            $this.gpoPermissions.Add(([GPOPermissions]::new($gpoName,$group,$permission)))
                        }
                    }
                }
            }
        } Else {
            # There are no GPO Permissiosn to process
            $this.ozoLogger.Write("No GPO permissions to process.","Warning")
        }
    }
    # METHODS: GrantDFSNRootPermissions method
    Hidden [Void] GrantDFSNRootPermissions() {
        # Determine if there are DFSN Root Permissions to process
        If (($this.Json.ADDFSNRootPermissions).Count -gt 0) {
            # There are DFSN Root Permissions to process; report
            $this.ozoLogger.Write("Processing DFSN root permissions.","Information")
            # Iterate through the DFSN Root Permissions
            ForEach ($dfsnRootPermission in $this.Json.ADDFSNRootPermissions) {
                # Report
                $this.ozoLogger.Write(("Processing " + $dfsnRootPermission.Description + "."),"Information")
                # Iterate through the Roots
                ForEach ($dfsnRoot in $dfsnRootPermission.DFSNRoots) {
                    # Iterate through the Identities
                    ForEach ($identity in $dfsnRootPermission.Identities) {
                        # Create a DFSNRootPermissions object for this Root + Identity
                        $this.dfsnRootPermissions.Add(([DFSNRootPermissions]::new($dfsnRoot,$identity)))
                    }
                }
            }
        } Else {
            # There are no DFSN Root Permissions to process
            $this.ozoLogger.Write("No DFSN root permissions to set.","Warning")
        }
    }
    # METHODS: GrantDFSNFolderPermissions method
    Hidden [Void] GrantDFSNFolderPermissions() {
        # Determine if there are DFSN Folder Permissions to process
        If (($this.Json.ADDFSNFolderPermissions).Count -gt 0) {
            # There are DFSN Folder Permissions to process
            $this.ozoLogger.Write("Processing DFSN folder permissions.","Information")
            # Iterate through the DFSN Folder Permissions
            ForEach ($dfsnFolderPermission in $this.Json.ADDFSNFolderPermissions) {
                # Report
                $this.ozoLogger.Write(("Processing " + $dfsnFolderPermission.Description + "."),"Information")
                # Iterate through the Folders
                ForEach ($dfsnFolder in $dfsnFolderPermission.DFSNFolders) {
                    # Iterate through the Identities
                    ForEach ($identity in $dfsnFolderPermission.Identities) {
                        # Create a DFSNFolderPermissions object for this Folder + Identity
                        $this.dfsnFolderPermissions.Add(([DFSNFolderPermissions]::new($dfsnFolder,$identity)))
                    }
                }
            }
        } Else {
            # There are no DFSN Folder Permissions to process
            $this.ozoLogger.Write("No DFSN folder permissions to set.","Warning")
        }
    }
    # METHODS: GrantDFSRPermissions method
    Hidden [Void] GrantDFSRPermissions() {
        # Determine if there are DFSR Permissions to process
        If (($this.Json.ADDFSRPermissions).Count -gt 0) {
            # There are DFSR Permissions to process; report
            $this.ozoLogger.Write("Processing DFSR Permissions.","Information")
            # Iterate through the DFSR Permissions
            ForEach ($dfsrPermission in $this.Json.ADDFSRPermissions) {
                # Report
                $this.ozoLogger.Write(("Processing " + $dfsrPermission.Description + "."),"Information")
                # Iterate through the Groups
                ForEach ($dfsrGroup in $dfsrPermission.DFSRGroups) {
                    # Iterate through the Identities
                    ForEach ($identity in $dfsrPermission.Identities) {
                        # Create a DFSRPermissions object for this Group + Identity
                        $this.dfsrPermissions.Add(([DFSRPermissions]::new($dfsrGroup,$identity)))
                    }
                }
            }
        } Else {
            # There are no DFSR Permissiosn to process
            $this.ozoLogger.Write("No DFSR permissions to set.","Warning")
        }
    }
    # Report method
    Hidden [Void] Report() {
        # Determine that at least one object was processed
        If (($this.ouDelegations + $this.gpoPermissions + $this.dfsnRootPermissions + $this.dfsnFolderPermissions + $this.dfsrPermissions).Count -gt 0) {
            # At least one object was processed; Produce Excel output
            $this.ouDelegations | Select-Object -Property @{Name="Delegation";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "Delegations" -Path $this.excelPath
            $this.gpoPermissions | Select-Object -Property @{Name="GPO Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "GPO Permissions" -Path $this.excelPath
            $this.dfsnRootPermissions | Select-Object -Property @{Name="DFSN Root Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "DFSN Root Permissions" -Path $this.excelPath
            $this.dfsnFolderPermissions | Select-Object -Property @{Name="DFSN Folder Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "DFSN Folder Permissions" -Path $this.excelPath
            $this.dfsrPermissions | Select-Object -Property @{Name="DFSR Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}},@{Name="Messages";Expression={$_.messages -Join "; "}} | Export-Excel -WorksheetName "DFSR Permissions" -Path $this.excelPath
            # Determine if session is interactive
            If ([Environment]::UserInteractive -eq $true) {
                # Session is interactive; produce output for the operator
                $this.ouDelegations | Select-Object -Property @{Name="Delegation";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}} | Format-Table | Out-Host
                $this.gpoPermissions | Select-Object -Property @{Name="GPO Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}} | Format-Table | Out-Host
                $this.dfsnRootPermissions | Select-Object -Property @{Name="DFSN Root Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}} | Format-Table | Out-Host
                $this.dfsnFolderPermissions | Select-Object -Property @{Name="DFSN Folder Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}} | Format-Table | Out-Host
                $this.dfsrPermissions | Select-Object -Property @{Name="DFSR Permission";Expression={$_.Description}},@{Name="Success";Expression={$_.Success}} | Format-Table | Out-Host
                # Determine if Excel was created
                Try {
                    Test-Path -Path $this.excelPath -ErrorAction Stop
                    # Success
                    $this.ozoLogger.Write(("For additional information, please see " + $this.excelPath + "."),"Information")
                } Catch {
                    $this.ozoLogger.Write("No Excel report generated.","Warning")
                }
            }
        } Else {
            # No objects were processed
            $this.ozoLogger.Write("No objects were processed.","Warning")
        }
    }
}

Class OUDelegation {
    # PROPERTIES: Booleans, Hashtables, Strings
    [Boolean]   $Success     = $false
    [Boolean]   $Validates   = $false
    [Hashtable] $erMap       = @{}
    [Hashtable] $guidMap     = @{}
    [String]    $dcOuDN      = $null
    [String]    $duOuDN      = $null
    [String]    $ouDN        = $null
    [String]    $Description = $null
    [String]    $Identity    = $null
    [String]    $identityDN  = $null
    [String]    $Permission  = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $Acl         = $null
    [PSCustomObject] $Permissions = @{
        CreateChildComputers = "Delegates create child computer objects"
        CreateChildContacts = "Delegates create child contacts objects"
        CreateChildGroups = "Delegates create child groups objects"
        CreateChildUsers = "Delegates create child user objects"
        CreateOUs = "Delegates create OU"
        DeleteChildComputers = "Delegates delete child computer objects"
        DeleteChildContacts = "Delegates delete child contacts objects"
        DeleteChildGroups = "Delegates delete child groups objects"
        DeleteChildUsers = "Delegates delete child user objects"
        DeleteOUs = "Delegates delete OU"
        DomainJoinComputer = "Delegates create computer objects, Write Name, and Write name"
        DomainLeaveComputer = "Delegates delete computer objects, Write Name, and Write name"
        EnableDisableComputers = "Delegates enable and disable computer objects"
        EnableDisableUsers = "Delegates enable and disable user objects"
        FullControlComputers = "Delegates full control to computer objects"
        FullControlContacts = "Delegates full control to contacts objects"
        FullControlGroups = "Delegates full control to group objects"
        FullControlUsers = "Delegates full control to user objects"
        FullControlOUs = "Delegates full control for organizational unit objects"
        LinkGPO = "Delegates link GPO"
        ModifyGroupMembership = "Delegates modify group membership"
        ResetUserPasswords = "Delegates reset password"
        ReadBitLockerRecovery = "Delegates read to the BitLocker Recovery information"
    }
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    OUDelegation($ouDN,$Identity,$Permission) {
        # Set properties
        $this.ouDN        = $ouDN
        $this.Identity    = $Identity
        $this.Permission  = $Permission
        $this.Description = ($this.Permission + " for " + $this.Identity + " on " + $this.ouDN)
        # Create a guid map and extended rights map
        $this.guidMap = (New-OZOADGuidMap)
        $this.erMap   = (New-OZOADExtendedRightsMap)
        # Determine if delegation validates
        If ($this.ValidateDelegation() -eq $true) {
            # Delegation validated
            $this.Validates = $true
            # Switch on permission
            Switch ($this.Permission) {
                "CreateChildComputers"   { $this.Success = $this.CreateChildComputers()   }
                "CreateChildContacts"    { $this.Success = $this.CreateChildContacts()    }
                "CreateChildGroups"      { $this.Success = $this.CreateChildGroups()      }
                "CreateChildUsers"       { $this.Success = $this.CreateChildUsers()       }
                "CreateOUs"              { $this.Success = $this.CreateOUs()              }
                "DeleteChildComputers"   { $this.Success = $this.DeleteChildComputers()   }
                "DeleteChildContacts"    { $this.Success = $this.DeleteChildContacts()    }
                "DeleteChildGroups"      { $this.Success = $this.DeleteChildGroups()      }
                "DeleteChildUsers"       { $this.Success = $this.DeleteChildUsers()       }
                "DeleteOUs"              { $this.Success = $this.DeleteOUs()              }
                "DomainJoinComputer"     { $this.Success = $this.DomainJoinComputer()     }
                "EnableDisableComputers" { $this.Success = $this.EnableDisableComputers() }
                "EnableDisableUsers"     { $this.Success = $this.EnableDisableUsers()     }
                "FullControlComputers"   { $this.Success = $this.FullControlComputers()   }
                "FullControlContacts"    { $this.Success = $this.FullControlContacts()    }
                "FullControlGroups"      { $this.Success = $this.FullControlGroups()      }
                "FullControlUsers"       { $this.Success = $this.FullControlUsers()       }
                "FullControlOUs"         { $this.Success = $this.FullControlOUs()         }
                "LinkGPO"                { $this.Success = $this.LinkGPO()                }
                "ModifyGroupMembership"  { $this.Success = $this.ModifyGroupMembership()  }
                "ReadBitLockerRecovery"  { $this.Success = $this.ReadBitLockerRecovery()  }
                "ResetUserPasswords"     { $this.Success = $this.ResetUserPasswords()     }
                default {
                    $this.Messages.Add(("No method found that matches " + $this.permission))
                    $this.Success = $false
                }
            }
        } Else {
            # Delegation did not validate
            $this.Messages.Add("Delegation does not validate")
            $this.Validates = $false
        }
    }
    # METHODS: ValidateDelegation method
    Hidden [Boolean] ValidateDelegation() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if ouDN is empty or null
        If ([String]::IsNullOrEmpty($this.ouDN)) {
            # ouDN is empty or null
            $this.Messages.Add("OU cannot be empty or null")
            $Return = $false
        }
        # Determine if Identity is empty or null
        If ([String]::IsNullOrEmpty($this.Identity)) {
            # Identity is empty or null
            $this.Messages.Add("Identity cannot be empty or null")
            $Return = $false
        } Else {
            # Try to get the distinguished name of the Identity
            Try {
                $this.identityDN = (Get-ADObject -Filter {SamAccountName -eq $this.Identity} -ErrorAction Stop).DistinguishedName
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("Identity does not exist in AD")
                $Return = $false
            }
        }
        # Determine if Permission is null or empty
        If ([String]::IsNullOrEmpty($this.Permission)) {
            # Permission is null or empty
            $this.Messages.Add("Permission cannot be empty or null")
            $Return = $false
        } Else {
            # Permission is not null or empty; determine if it does not match the permissions hashtable
            If (($this.Permissions.PSObject.Properties -Match $this.Permission) -eq $false) {
                # Permission does not appear in the hash table
                $this.Messages.Add("Permission does not exist")
                $Return = $false
            }
        }
        # Try to get the ACL for the OU
        Try {
            $this.Acl = Get-ACL ("AD:" + $this.ouDN) -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Unable to obtain ACL for " + $this.ouDN))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: CreateChildUsers method
    Hidden [Boolean] CreateChildComputers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create child computers delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: CreateChildContacts method
    Hidden [Boolean] CreateChildContacts() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Contact -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create child contacts delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: CreateChildGroups method
    Hidden [Boolean] CreateChildGroups() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Group -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create child groups delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: CreateChildUsers method
    Hidden [Boolean] CreateChildUsers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create child users delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: CreateDeleteOUs method
    Hidden [Boolean] CreateOUs() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate create OU
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeGuid ($this.guidMap['OrganizationalUnit']) -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create OU delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: DeleteChildUsers method
    Hidden [Boolean] DeleteChildComputers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding delete child computers delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: DeleteChildContacts method
    Hidden [Boolean] DeleteChildContacts() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Contact -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding delete child contacts delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: DeleteChildGroups method
    Hidden [Boolean] DeleteChildGroups() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Group -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding delete child groups delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: DeleteChildUsers method
    Hidden [Boolean] DeleteChildUsers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding delete child users delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: CreateDeleteOUs method
    Hidden [Boolean] DeleteOUs() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate delete  OU
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeGuid ($this.guidMap['OrganizationalUnit']) -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding delete OU delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: DomainJoinComputer method
    Hidden [Boolean] DomainJoinComputer() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLJoinDomain -TargetDN $this.ouDN -DelegateDN $this.identityDN -AllowCreate -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding join domain delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: EnableDisableComputers method
    Hidden [Boolean] EnableDisableComputers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            # Add-DSACLWriteAccountRestrictions -TargetDN <String> -DelegateDN <String> -ObjectTypeName <String> -AccessType <AccessControlType> [-NoInheritance] [<CommonParameters>]
            Add-DSACLWriteAccountRestrictions -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding enable disable users delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: EnableDisableUsers method
    Hidden [Boolean] EnableDisableUsers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            # Add-DSACLWriteAccountRestrictions -TargetDN <String> -DelegateDN <String> -ObjectTypeName <String> -AccessType <AccessControlType> [-NoInheritance] [<CommonParameters>]
            Add-DSACLWriteAccountRestrictions -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding enable disable users delegation. Error message is " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: FullControlComputers method
    Hidden [Boolean] FullControlComputers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLFullControl -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding full control computers delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: FullControlContacts method
    Hidden [Boolean] FullControlContacts() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLFullControl -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Contact -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding full control contacts delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: FullControlGroup method
    Hidden [Boolean] FullControlGroups() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLFullControl -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Group -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding full control groups delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: FullControlUsers method
    Hidden [Boolean] FullControlUsers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLFullControl -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding full control users delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: FullControlUsers method
    Hidden [Boolean] FullControlOUs() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLFullControl -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeGuid ($this.guidMap['OrganizationalUnit']) -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding full control organizational units delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: LinkGPO method
    Hidden [Boolean] LinkGPO() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLLinkGPO -TargetDN $this.ouDN -DelegateDN $this.identityDN -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding link GPO delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: ModifyGroupMembership method
    Hidden [Boolean] ModifyGroupMembership() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLManageGroupMember -TargetDN $this.ouDN -DelegateDN $this.identityDN -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding modify group membership delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: ReadBitLockerRecovery method
    Hidden [Boolean] ReadBitLockerRecovery() {
        # Control variable
        [Boolean] $Return = $true
        # Set parameters
        [Hashtable] $Parameters = @{
            TargetDN              = $this.ouDN
            DelegatDN             = $this.identityDN
            ActiveDirectoryRights = [System.DirectoryServices.ActiveDirectoryRights]"ReadProperty,ExtendedRight"
            AccessControlType     = [System.Security.AccessControl.AccessControlType]"Allow"
            ObjectType            = [Guid]"00000000-0000-0000-0000-000000000000"
            InheritanceType       = [System.DirectoryServices.ActiveDirectorySecurityInheritance]"All"
            InheritedObjectType   = [Guid]$this.guidMap["msFVE-RecoveryInformation"]
        }
        # Try to delegate
        Try {
            Add-DSACLCustom @Parameters -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding read bitlocker recovery information delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # METHODS: Reset password method
    Hidden [Boolean] ResetUserPasswords() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate
        Try {
            Add-DSACLResetPassword -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding reset password delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
}

Class GPOPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $gpoName     = $null
    [String]  $Group       = $null
    [String]  $Permission  = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $Permissions = [PSCustomObject]@{
        GpoRead  = "Allows reading a named GPO"
        GpoApply = "placeholder"
        GpoEdit  = "Allows editing a named GPO"
        GpoEditDeleteModifySecurity = "placeholder"
    }
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    GPOPermissions($gpoName,$group,$permission) {
        # Set properties
        $this.gpoName     = $gpoName
        $this.Group       = $group
        $this.Permission  = $permission
        $this.Description = ($this.Permission + " for " + $Group + " on " + $gpoName)
        # Determine if the GPO permissions validate
        If ($this.ValidateGPOPermission() -eq $true) {
            # GPO permissions validate
            $this.Validates = $true
            $this.Success   = $this.SetGPOPermission()
        } Else {
            # GPO permissions do not validate
            $this.Messages.Add("GPO permission does not validate")
            $this.Validates = $false
        }
    }
    # METHODS: ValidateGPOPermission method
    Hidden [Boolean] ValidateGPOPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if gpoName is null or empty
        If ([String]::IsNullOrEmpty($this.gpoName)) {
            # gpoName is null or empty
            $this.Messages.Add("GpoName cannot be empty or null")
            $Return = $false
        } Else {
            # gpoName is not null or empty; try to get the GPO
            Try {
                Get-GPO -Name $this.gpoName -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("GPO does not exist in AD")
                $Return = $false
            }
        }
        # Determine if Group is null or empty
        If ([String]::IsNullOrEmpty($this.Group)) {
            # Group is null or empty
            $this.Messages.Add("Group cannot be empty or null")
            $Return = $false
        } Else {
            # Group is not null or empty; try to get the group
            Try {
                Get-ADObject -Filter {SamAccountName -eq $this.Group} -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("Group does not exist in AD")
                $Return = $false
            }
        }
        # Determine if Permission is null or empty
        If ([String]::IsNullOrEmpty($this.Permission)) {
            # Permission is null or empty
            $this.Messages.Add("Permission cannot be empty or null")
            $Return = $false
        } Else {
            # Permission is not null or empty; detemine if Permissions exists in the Permissions object
            If (($this.Permissions.PSObject.Properties -Match $this.Permission) -eq $false) {
                # Permission does not exist in Permisssions object
                $this.Messages.Add("Permission is not handled")
                $Return = $false
            }
        }
        # Return
        return $Return
    }
    # METHODS: GpoEdit method
    Hidden [Boolean] SetGPOPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Try to set GPO permission
        Try {
            Set-GPPermission -Name $this.gpoName -TargetName $this.Group -TargetType Group -PermissionLevel $this.Permission -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add("Error setting GPO permission")
            $Return = $false
        }
        # Return
        return $Return
    }
}

Class DFSNRootPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $dfsnRoot    = $null
    [String]  $Identity    = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    DFSNRootPermissions($dfsnRoot,$identity) {
        # Set properties
        $this.dfsnRoot = $dfsnRoot
        $this.Identity = $identity
        $this.Description = ("Permission for " + $this.Identity + " on " + $this.dfsnRoot)
        # Determine if the DFSN root permission validates
        If ($this.ValidateDFSNRootPermission() -eq $true) {
            # DFSN root permission validates
            $this.Validates = $true
            # Call SetDFSNRootPermission to set Success
            $this.Success = $this.SetDFSNRootPermission()
        } Else {
            # DFSN root permission does not validate
            $this.Messages.Add("DFSN root permission does not validate.")
            $this.Validates = $false
        }
    }
    # METHODS: ValidateDFSNPermission method
    Hidden [Boolean] ValidateDFSNRootPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if dfsnRoot is null or empty
        If ([String]::IsNullOrEmpty($this.dfsnRoot)) {
            # dfsnRoot is null or empty
            $this.Messages.Add("DFSNRoot cannot be empty or null")
            $Return = $false
        } Else {
            # dfsnRoot is not null or empty; try to get the DFSN root
            Try {
                Get-DfsnRoot -Path $this.dfsnRoot -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("DFSN root does not exist in DFS. Error message is: " + $_))
                $Return = $false
            }
        }
        # Determine if Identity is null or empty
        If ([String]::IsNullOrEmpty($this.Identity)) {
            # Identity is null or empty
            $this.Messages.Add("Identity cannot be empty or null")
            $Return = $false
        } Else {
            # Identity is not null or empty; try to get the object
            Try {
                Get-ADObject -Filter {SamAccountName -eq $this.Identity} -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("Identity does not exist in AD")
                $Return = $false
            }
        }
        # Return
        return $Return
    }
    # METHODS: SetDFSNPermission
    Hidden [Boolean] SetDFSNRootPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Try to set the DFSN root permission
        Try {
            Set-DfsnRoot -GrantAdminAccounts $this.Identity -Path $this.dfsnRoot -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error setting DFSN root permission. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }   
}

Class DFSNFolderPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $dfsnFolder  = $null
    [String]  $Identity    = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    DFSNFolderPermissions($dfsnFolder,$identity) {
        # Set properties
        $this.dfsnFolder  = $dfsnFolder
        $this.Identity    = $identity
        $this.Description = ("Permission for " + $this.Identity + " on " + $this.dfsnFolder)
        # Determine if the DFSN folder permission validates
        If ($this.ValidateDFSNFolderPermission() -eq $true) {
            # DFSN folder permission validates
            $this.Validates = $true
            # Call SetDFSNFolderPermissions to set Success
            $this.Success = $this.SetDFSNFolderPermission()
        } Else {
            # DFSN folder permission does not validate
            $this.Messages.Add(("DFSN folder permission " + $this.Description + " does not validate"))
            $this.Validates = $false
        }
    }
    # METHODS: ValidateDFSNPermission method
    Hidden [Boolean] ValidateDFSNFolderPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if dfsnFolder is null or empty
        If ([String]::IsNullOrEmpty($this.dfsnFolder)) {
            # dfsnFolder is null or ermpty
            $this.Messages.Add("DFSNFolder cannot be empty or null")
            $Return = $false
        } Else {
            # dfsnFolder is not null or empty; attempt to get the folder
            Try {
                Get-DfsnFolder -Path $this.dfsnFolder -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("DFSN folder does not exist in DFS. Error message is: " + $_))
                $Return = $false
            }
        }
        # Determine if Identity is null or empty
        If ([String]::IsNullOrEmpty($this.Identity)) {
            # Identity is null or empty
            $this.Messages.Add("Identity cannot be empty or null")
            $Return = $false
        } Else {
            # Identity is not null or empty
            Try {
                Get-ADObject -Filter {SamAccountName -eq $this.identity} -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("Identity does not exist in AD")
                $Return = $false
            }
        }
        # Return
        return $Return
    }
    # METHODS: SetDFSNPermission
    Hidden [Boolean] SetDFSNFolderPermission() {
        # Control variable
        $Return = $true
        # Try to grant the DFSN folder permission
        Try {
            Grant-DfsnAccess -Path $this.dfsnFolder -AccountName $this.Identity -ErrorAction Stop
            # Success
        } Catch {
            $this.Messages.Add(("Error setting DFSN folder permission. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }   
}

Class DFSRPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $dfsrGroup   = $null
    [String]  $Identity    = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS: Constructor method
    DFSRPermissions($dfsrGroup,$identity) {
        $this.dfsrGroup   = $dfsrGroup
        $this.identity    = $identity
        $this.Description = ("Permission for " + $this.Identity + " to " + $this.dfsrGroup)
        # Determine if the DFSR permission validates
        If ($this.ValidateDFSRPermission() -eq $true) {
            # DFSR permission validates
            $this.Validates = $true
            # Call SetDFSRPermission to set Success
            $this.Success = $this.SetDFSRPermission()
        } Else {
            # DFSR permission does not validate
            $this.Messages.Add("Permission does not validate")
            $this.Validates = $false
        }
    }
    # METHODS: ValidateDFSRPermission method
    Hidden [Boolean] ValidateDFSRPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if dfsrGroup is null or empty
        If ([String]::IsNullOrEmpty($this.dfsrGroup)) {
            # dfsrGroup is null or empty
            $this.Messages.Add("DFSRGroup cannot be empty or null")
            $Return = $false
        } Else {
            # dfsrGroup is not null or empty; try to get replication grpu
            Try {
                Get-DfsReplicationGroup -GroupName $this.dfsrGroup -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("DFSR group does not exist in DFS")
                $Return = $false
            }
        }
        # Determine if Identity is null or empty
        If ([String]::IsNullOrEmpty($this.Identity)) {
            # Identity is null or empty
            $this.Messages.Add("Identity cannot be empty or null")
            $Return = $false
        } Else {
            # Identity is not null or empty; attempt to get the object
            Try {
                Get-ADObject -Filter {SamAccountName -eq $this.Identity} -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add("Identity does not exist in AD")
                $Return = $false
            }
        }
        return $Return
    }
    # METHODS: SetDFSRPermission method
    Hidden [Boolean] SetDFSRPermission() {
        # Control variable
        [Boolean] $Return = $true
        # Try to grant permission
        Try {
            Grant-DfsrDelegation -GroupName $this.dfsrGroup -AccountName $this.identity -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add("Error setting DFSR permission")
            $Return = $false
        }
        # Return
        return $Return
    }   
}

# MAIN
[Main]::new($Configuration,$OutDir) | Out-Null
