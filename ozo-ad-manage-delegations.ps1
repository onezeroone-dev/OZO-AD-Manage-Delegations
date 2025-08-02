#Requires -Modules ActiveDirectory,DFSN,DFSR,DSACL,GroupPolicy,ImportExcel,SH,SHAD

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
    .RELEASENOTES
#>

<# 
    .SYNOPSIS
    See description.
    .DESCRIPTION 
    Creates AD delegations based on a configuration file. This script can create OU Delegations (`OUDelegations`), apply permissions to GPOs (`GPOPermissions`), grant access to DFSN roots (`DFSNRootPermissions`), grant access to DFSN folders (`DFSNFolderPermissions`) and create delegations to DFSR replication groups ("DFSRPermissions").
    .PARAMETER Configuration
    Path to the JSON configuration file. Defaults to "ad-create-delegations.json" in the same directory as the script.
    .PARAMETER OutDir
    Path for the Excel report. Defaults to the current directory.
    .PARAMETER Wipe
    PENDING IMPLEMENTATION wipe all existing delegations before applying the configured delegations.
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
    Run this script as a user with rights to create AD delegations (likely a Domain Admin) from within a writable directory.
#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess = $true)] Param (
    [Parameter(Mandatory=$false,HelpMessage="Path to the JSON configuration file")][String]$Configuration = (Join-Path -Path $PSScriptRoot -ChildPath "ad-create-delegations.json"),
    [Parameter(Mandatory=$false,HelpMessage="Path for the Excel report")][String]$OutDir = (Get-Location),
    [Parameter(Mandatory=$false,HelpMessage="PENDING IMPLEMENTATION wipe all existing delegations before applying the configured delegations")][Switch]$Wipe
)

# CLASSES
Class ACDMain {
    # PROPERTIES: Booleans, Hashtables, Strings
    [Boolean]   $Validates = $true
    [Boolean]   $Wipe      = $false
    [String]    $excelPath = $null
    [String]    $jsonPath  = $null
    [String]    $outDir    = $null
    # PROPERTIES: PSCustomObjects
    [PSCustomObject] $Json     = $null
    [PSCustomObject] $shLogger = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[PSCustomObject]] $ouDelegations         = @()
    [System.Collections.Generic.List[PSCustomObject]] $gpoPermissions        = @()
    [System.Collections.Generic.List[PSCustomObject]] $dfsnRootPermissions   = @()
    [System.Collections.Generic.List[PSCustomObject]] $dfsnFolderPermissions = @()
    [System.Collections.Generic.List[PSCustomObject]] $dfsrPermissions       = @()
    # METHODS
    # Constructor method
    ACDMain($Configuration,$OutDir,$Wipe) {
        # Set properties
        $this.jsonPath = $Configuration
        $this.outDir   = $OutDir
        $this.Wipe     = $Wipe
        # Create a shLogger object
        $this.shLogger = (New-SHLogger)
        # Declare ourselves to the world
        $this.shLogger.Log("Starting process.","Information")
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
        $this.shLogger.Log("Process complete.","Information")
    }
    # Configuration validation method
    Hidden [Boolean] ValidateConfiguration() {
        # control variable
        [Boolean]$Return = $true
        # Check that the jsonPath is valid
        Try {
            Test-Path -Path $this.jsonPath -ErrorAction Stop
            # Success; attempt to read the JSON
            Try {
                $this.json = Get-Content $this.jsonPath -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                # Success (able to read JSON)
                $this.shLogger.Log("Configuration validates.","Information")
            } Catch {
                # Failure (unable to read JSON)
                $this.shLogger.Log(("Invalid JSON in " + $this.jsonPath + "."),"Error")
                $Return = $false
            }
        } Catch {
            # JSON path is not valid
            $this.shLogger.Log(("Could not read configuration file " + $this.jsonPath + "."),"Error")
            $Return = $false
        }
        # Return
        return $Return
    }
    # Environment validation method
    Hidden [Boolean] ValidateEnvironment() {
        # Control variable
        [Boolean] $Return = $true
        # Detemine if session is not user-interactive
        If ([Environment]::UserInteractive -eq $false) {
            $this.shLogger.Log("Please run this script in an interactive session.","Error")
            $Return = $false
        }
        # Determine if outDir is writable
        If ((Test-SHPathWritable -Path $this.outDir) -eq $true) {
            # outDir is writable
            $this.excelPath = (Join-Path -Path $this.outDir -ChildPath ((Get-SH8601Date -Time) + "-ad-create-delegations-report.xlsx"))
            $this.shLogger.Log(("Using " + $this.excelPath + " for the Excel report"),"Information")
        } Else {
            # outDir is not writable; determine if current location is writable
            $this.shLogger.Log("Provided output directory is not writable.","Warning")
            If ((Test-SHPathWritable -Path (Get-Location)) -eq $true) {
                # current directory is writable
                $this.excelPath = (Join-Path -Path (Get-Location) -ChildPath ((Get-SH8601Date -Time) + "-ad-create-delegations-report.xlsx"))
                $this.shLogger.Log(("Using " + $this.excelPath + " for the Excel report"),"Information")
            } Else {
                # current directory is not writable
                $this.shLogger.Log("Current directory is not writable; cannot proceed.","Error")
                $return = $false
            }
        }
        # Return
        return $Return
    }
    # CreateDelegations method
    Hidden [Void] CreateOUDelegations() {
        If (($this.json.OUDelegations).Count -gt 0) {
            $this.shLogger.Log("Processing Delegations.","Information")
            ForEach ($delegation in $this.json.OUDelegations) {
                $this.shLogger.Log(("Processing " + $delegation.Description + "."),"Information")
                ForEach ($ouDN in $delegation.OUs) {
                    ForEach ($identity in $delegation.Identities) {
                        ForEach ($permission in $delegation.Permissions) {
                            $this.ouDelegations.Add(([ACDOUDelegation]::new($ouDN,$this.json.DisabledComputersOUDN,$this.json.DisabledUsersOUDN,$identity,$permission)))
                        }
                    }
                }
            }
        } Else {
            $this.shLogger.Log("No delegations to process.","Warning")
        }
    }
    # SetGPOPermissions method
    Hidden [Void] SetGPOPermissions() {
        If (($this.json.GPOPermissions).Count -gt 0) {
            $this.shLogger.Log("Processing GPO Permissions.","Information")
            ForEach ($gpoPermission in $this.json.GPOPermissions) {
                $this.shLogger.Log(("Processing " + $gpoPermission.Description + "."),"Information")
                ForEach ($gpoName in $gpoPermission.GPONames) {
                    ForEach ($group in $gpoPermission.GroupNames) {
                        ForEach ($permission in $gpoPermission.Permissions) {
                            $this.gpoPermissions.Add(([ACDGPOPermissions]::new($gpoName,$group,$permission)))
                        }
                    }
                }
            }
        } Else {
            $this.shLogger.Log("No GPO permissions to process.","Warning")
        }
    }
    # GrantDFSNRootPermissions method
    Hidden [Void] GrantDFSNRootPermissions() {
        If (($this.json.DFSNRootPermissions).Count -gt 0) {
            $this.shLogger.Log("Processing DFSN Permissions.","Information")
            ForEach ($dfsnRootPermission in $this.json.DFSNRootPermissions) {
                $this.shLogger.Log(("Processing " + $dfsnRootPermission.Description + "."),"Information")
                ForEach ($dfsnRoot in $dfsnRootPermission.DFSNRoots) {
                    ForEach ($identity in $dfsnRootPermission.Identities) {
                        $this.dfsnRootPermissions.Add(([ACDDFSNRootPermissions]::new($dfsnRoot,$identity)))
                    }
                }
            }
        } Else {
            $this.shLogger.Log("No DFSN permissions to set.","Warning")
        }
    }
    # GrantDFSNFolderPermissions method
    Hidden [Void] GrantDFSNFolderPermissions() {
        If (($this.json.DFSNFolderPermissions).Count -gt 0) {
            $this.shLogger.Log("Processing DFSN Permissions.","Information")
            ForEach ($dfsnFolderPermission in $this.json.DFSNFolderPermissions) {
                $this.shLogger.Log(("Processing " + $dfsnFolderPermission.Description + "."),"Information")
                ForEach ($dfsnFolder in $dfsnFolderPermission.DFSNFolders) {
                    ForEach ($identity in $dfsnFolderPermission.Identities) {
                        $this.dfsnFolderPermissions.Add(([ACDDFSNFolderPermissions]::new($dfsnFolder,$identity)))
                    }
                }
            }
        } Else {
            $this.shLogger.Log("No DFSN permissions to set.","Warning")
        }
    }
    # GrantDFSRPermissions method
    Hidden [Void] GrantDFSRPermissions() {
        If (($this.json.DFSRPermissions).Count -gt 0) {
            $this.shLogger.Log("Processing DFSR Permissions.","Information")
            ForEach ($dfsrPermission in $this.json.DFSRPermissions) {
                $this.shLogger.Log(("Processing " + $dfsrPermission.Description + "."),"Information")
                ForEach ($dfsrGroup in $dfsrPermission.DFSRGroups) {
                    ForEach ($identity in $dfsrPermission.Identities) {
                        $this.dfsrPermissions.Add(([ACDDFSRPermissions]::new($dfsrGroup,$identity)))
                    }
                }
            }
        } Else {
            $this.shLogger.Log("No DFSR permissions to set.","Warning")
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
                    $this.shLogger.Log(("For additional information, please see " + $this.excelPath + "."),"Information")
                } Catch {
                    $this.shLogger.Log("No Excel report generated.","Warning")
                }
            }
        } Else {
            # No objects were processed
            $this.shLogger.Log("No objects were processed.","Warning")
        }
    }
}

Class ACDOUDelegation {
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
        CreateDeleteComputers = "Delegates create computer and delete computer"
        CreateDeleteOUs = "Delegates create OU and delete OU"
        CreateDeleteUsers = "Delegates create user and delete user"
        DomainJoinComputer = "Delegates create computer objects, Write Name, and Write name"
        DomainLeaveComputer = "Delegates delete computer objects, Write Name, and Write name"
        EnableDisableComputers = "Delegates enable and disable computer objects"
        EnableDisableUsers = "Delegates enable and disable user objects"
        FullControlComputers = "Delegates full control to computer objects"
        FullControlContacts = "Delegates full control to contacts objects"
        FullControlGroups = "Delegates full control to group objects"
        FullControlUsers = "Delegates full control to user objects"
        LinkGPO = "Delegates link GPO"
        ModifyGroupMembership = "Delegates modify group membership"
        MoveUsersToDisabledUsersOU = "Delegates moving a disabled user to the DisabledUsers OU"
        MoveComputersToDisabledComputersOU = "Delegates moving a disabled computer to the DisabledComputers OU"
        ResetUserPasswords = "Delegates reset password"
        ReadBitLockerRecovery = "Delegates read to the BitLocker Recovery information"
    }
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS
    # Constructor method
    ACDOUDelegation($ouDN,$dcOuDN,$duOuDN,$Identity,$Permission) {
        # Set properties
        $this.ouDN        = $ouDN
        $this.dcOuDN      = $dcOuDN
        $this.duOuDN      = $duOuDN
        $this.Identity    = $Identity
        $this.Permission  = $Permission
        $this.Description = ($this.Permission + " for " + $this.Identity + " on " + $this.ouDN)
        # Create a guid map and extended rights map
        $this.guidMap = (New-SHADGuidMap)
        $this.erMap   = (New-SHADExtendedRightMap)
        # Determine if delegation validates
        If ($this.ValidateDelegation() -eq $true) {
            # Delegation validated
            $this.Validates = $true
            # Switch on permission
            Switch ($this.Permission) {
                "CreateChildComputers"               { $this.Success = $this.CreateChildComputers()               }
                "CreateChildContacts"                { $this.Success = $this.CreateChildContacts()                }
                "CreateChildGroups"                  { $this.Success = $this.CreateChildGroups()                  }
                "CreateChildUsers"                   { $this.Success = $this.CreateChildUsers()                   }
                "CreateDeleteComputers"              { $this.Success = $this.CreateDeleteComputers()              }
                "CreateDeleteOUs"                    { $this.Success = $this.CreateDeleteOUs()                    }
                "CreateDeleteUsers"                  { $this.Success = $this.CreateDeleteUsers()                  }
                "DomainJoinComputer"                 { $this.Success = $this.DomainJoinComputer()                 }
                "DomainLeaveComputer"                { $this.Success = $this.DomainLeaveComputer()                }
                "EnableDisableComputers"             { $this.Success = $this.EnableDisableComputers()             }
                "EnableDisableUsers"                 { $this.Success = $this.EnableDisableUsers()                 }
                "FullControlComputers"               { $this.Success = $this.FullControlComputers()               }
                "FullControlContacts"                { $this.Success = $this.FullControlContacts()                }
                "FullControlGroups"                  { $this.Success = $this.FullControlGroups()                  }
                "FullControlUsers"                   { $this.Success = $this.FullControlUsers()                   }
                "LinkGPO"                            { $this.Success = $this.LinkGPO()                            }
                "ModifyGroupMembership"              { $this.Success = $this.ModifyGroupMembership()              }
                "MoveComputerObject"                 { $this.Success = $this.MoveComputerObject()                 }
                "MoveUsersToDisabledUsersOU"         { $this.Success = $this.MoveUsersToDisabledUsersOU()         }
                "MoveComputersToDisabledComputersOU" { $this.Success = $this.MoveComputersToDisabledComputersOU() }
                "ReadBitLockerRecovery"              { $this.Success = $this.ReadBitLockerRecovery()              }
                "ResetUserPasswords"                 { $this.Success = $this.ResetUserPasswords()                 }
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
    # ValidateDelegation method
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
    # CreateChildUsers method
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
    # CreateChildContacts method
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
    # CreateChildGroups method
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
    # CreateChildUsers method
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
    # CreateDeleteComputers method
    Hidden [Boolean] CreateDeleteComputers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate create computer
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create computer delegation. Error message is: " + $_))
            $Return = $false
        }
        # Try to delegate delete computer
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding delete computer delegation. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # CreateDeleteOUs method
    Hidden [Boolean] CreateDeleteOUs() {
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
    # CreateDeleteUsers method
    Hidden [Boolean] CreateDeleteUsers() {
        # Control variable
        [Boolean] $Return = $true
        # Try to delegate create user
        Try {
            Add-DSACLCreateChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            # Failure
            $this.Messages.Add(("Error adding create user delegation. Error message is: " + $_))
            $Return = $false
        }
        # Try to delegate delete user
        Try {
            Add-DSACLDeleteChild -TargetDN $this.ouDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
            # Success
        } Catch {
            $this.Messages.Add(("Error adding delete user delegation. Error message is: " + $_))
            # Failure
            $Return = $false
        }
        # Return
        return $Return
    }
    # DomainJoinComputer method
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
    # DomainLeaveComputer method
    Hidden [Boolean] DomainLeaveComputer() {
        # Control variable
        [Boolean] $Return = $false
        # Create computer objects, Write Name, Write name
        $this.Messages.Add("The DomainLeaveComputer permission method is not yet implemented.")
        # Return
        return $Return
    }
    # EnableDisableUsers method
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
    # EnableDisableUsers method
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
    # FullControlComputers method
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
    # FullControlContacts method
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
    # FullControlGroup method
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
    # FullControlUsers method
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
    # LinkGPO method
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
    # ModifyGroupMembership method
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
    # MoveComputerObject method
    Hidden [Boolean] MoveComputerObject() {
        # Control variable
        [Boolean] $Return = $false
        $this.Messages.Add("The MoveComputerObject permission method is not yet implemented. This probably needs to be a method in a different class.")
        # Return
        return $Return
    }
    # MoveDisabledUser method
    Hidden [Boolean] MoveUsersToDisabledUsersOU() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if we have been provided a disabled users OU
        If ([String]::IsNullOrEmpty($this.duOuDN) -eq $false) {
            # We have been provided a disabled users OU; attempt to add delegation.
            Try {
                # Per DSACL documentation, the identity responsible for moving user objects needs create child rights in the destination OU. In this case, "target" means the OU where the object will be moved *TO*
                Add-DSACLCreateChild -TargetDN $this.duOuDN -DelegateDN $this.identityDN -ObjectTypeName User -AccessType Allow -ErrorAction Stop
                # Per DSACL documentation, the identity responsible for moving user objects needs this rename and delete objects in the target OU. In this case, "target" means the OU where the object will be moved *FROM*
                Add-DSACLMoveObjectFrom -ObjectTypeName User -TargetDN $this.ouDN -DelegateDN $this.identityDN -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Error adding move user to diabled users OU delegation. Error message is: " + $_))
                $Return = $false
            }
        } Else {
            # We have not been provided a disabled users OU; skipping
            $this.Messages.Add(("No disabled users OU provided; skipping. Error message is: " + $_))
            $Return = $false
        }
        # Return
        return $Return
    }
    # MoveDisabledComputer method
    Hidden [Boolean] MoveComputersToDisabledComputersOU() {
        # Control variable
        [Boolean] $Return = $true
        # Determine if we have been provided a disabled computers OU
        If ([String]::IsNullOrEmpty($this.dcOuDN) -eq $false) {
            # We have been provided a disabled users OU; try to add delegation.
            Try {
                # Per DSACL documentation, the identity responsible for moving computer objects needs create child rights in the destination OU. In this case, "target" means the OU where the object will be moved *TO*
                Add-DSACLCreateChild -TargetDN $this.dcOuDN -DelegateDN $this.identityDN -ObjectTypeName Computer -AccessType Allow -ErrorAction Stop
                # Per DSACL documentation, the identity responsible for moving computer objects needs this rename and delete objects in the target OU. In this case, "target" means the OU where the object will be moved *FROM*
                Add-DSACLMoveObjectFrom -ObjectTypeName Computer -TargetDN $this.ouDN -DelegateDN $this.identityDN -ErrorAction Stop
                # Success
            } Catch {
                # Failure
                $this.Messages.Add(("Error adding move computer to diabled users OU delegation. Error message is: " + $_))
                $Return = $false
            }
        } Else {
            # We have not been provided a disabled users OU; skipping
            $this.Messages.Add("No disabled users OU provided; skipping")
            $Return = $false
        }
        # Return
        return $Return
    }
    # ReadBitLockerRecovery method
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
    # Reset password method
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

Class ACDGPOPermissions {
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
    # METHODS
    # Constructor method
    ACDGPOPermissions($gpoName,$group,$permission) {
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
    # ValidateGPOPermission method
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
    # GpoEdit method
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

Class ACDDFSNRootPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $dfsnRoot    = $null
    [String]  $Identity    = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS
    # Constructor method
    ACDDFSNRootPermissions($dfsnRoot,$identity) {
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
    # ValidateDFSNPermission method
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
    # SetDFSNPermission
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

Class ACDDFSNFolderPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $dfsnFolder  = $null
    [String]  $Identity    = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS
    # Constructor method
    ACDDFSNFolderPermissions($dfsnFolder,$identity) {
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
    # ValidateDFSNPermission method
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
    # SetDFSNPermission
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

Class ACDDFSRPermissions {
    # PROPERTIES: Booleans, Strings
    [Boolean] $Success     = $false
    [Boolean] $Validates   = $false
    [String]  $Description = $null
    [String]  $dfsrGroup   = $null
    [String]  $Identity    = $null
    # PROPERTIES: Lists
    [System.Collections.Generic.List[String]] $Messages = @()
    # METHODS
    # Constructor method
    ACDDFSRPermissions($dfsrGroup,$identity) {
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
    # ValidateDFSRPermission method
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
    # SetDFSRPermission method
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
[ACDMain]::new($Configuration,$OutDir,$Wipe) | Out-Null
