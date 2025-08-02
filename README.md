# OZO AD Manage Delegations Installation and Usage
## Description
Creates AD delegations based on a configuration file. This script can create OU Delegations (`OUDelegations`), apply permissions to GPOs (`GPOPermissions`), grant access to DFSN roots (`DFSNRootPermissions`), grant access to DFSN folders (`DFSNFolderPermissions`) and create delegations to DFSR replication groups (`DFSRPermissions`).

The JSON configuration file and resulting report can be useful as documentation and as evidence for a security audit.

## Prerequisites
This script requires the _ActiveDirectory_, _DFSN_, _DFSR_, _DSACL_, _GroupPolicy_, _ImportExcel_, _OZO_, and _OZOAD_ PowerShell modules. The ActiveDirectory, DFSN, DFSR, and GroupPolicy modules are included with the [Remote Server Administration Tools](https://learn.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools) installation. The remaining modules are published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Module DSACL,ImportExcel,OZO,OZOAD
```

## Installation
This script is published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Script ozo-ad-manage-delegations
```

## Usage
```
ozo-ad-manage-delegations
    [-Configuration <String>]
    [-OutDir        <String>]
```

## Parameters
|Parameter|Description|
|---------|-----------|
|`Configuration`|Path to the JSON configuration file. Defaults to `ozo-ad-manage-delegations.json` in the same directory as the script. Please see _Configuration Definition_ (below) for more information.|
|`OutDir`|Directory for the Excel report. Defaults to the current directory.|

## Capabilities
### OUDelegations
The following permissions are implemented in the script. See [_Configuration Definition_](#configuration-definition) below for information on how to use these permissions.

|Permission|Description|Implemented|
|----------|-----------|------|
|`CreateChildComputers`|Delegates create child computer objects.|TRUE|
|`CreateChildContacts`|Delegates create child contact objects.|TRUE|
|`CreateChildGroups`|Delegates create child group objects.|TRUE|
|`CreateChildUsers`|Delegates create child user objects.|TRUE|
|`CreateDeleteComputers`|Delegates create computer and delete computer.|TRUE|
|`CreateDeleteOUs`|Delegates create OU and delete OU.|TRUE|
|`CreateDeleteUsers`|Delegates create user and delete user.|TRUE|
|`DomainJoinComputer`|Delegates create computer objects, Write Name, and Write name.|TESTING|
|`DomainLeaveComputer`|Delegates delete computer objects, Write Name, and Write name.|FALSE|
|`EnableDisableUsers`|Delegates enable and disable user objects|TRUE|
|`FullControlComputers`|Delegates full control to computer objects.|TRUE|
|`FullControlContacts`|Delegates full control to contact objects.|TRUE|
|`FullControlGroups`|Delegates full control to group objects.|TRUE|
|`FullControlUsers`|Delegates full control to user objects.|TRUE|
|`LinkGPO`|Delegates permission to link GPOs.|TRUE|
|`ModifyGroupMembership`|Delegates modify group membership.|TRUE|
|`MoveUsersToDisabledUsersOU`|Delegates move user from Users OU to the Disabled Users OU.|TRUE|
|`ReadBitLockerRecovery`|Delegates read to the BitLocker Recovery information.|TRUE|
|`ResetUserPasswords`|Delegates reset password.|TRUE|

### GPOPermissions
The script can apply the Group Policy permissions `GpoRead`, `GpoApply`, `GpoEdit`, and `GpoEditDeleteModifySecurity` for _groups_.

### DFSNRootPermissions
The script can grant access to a DFSN **root** to a user or group.

### DFSNFolderPermissions
The script can grant access to DFSN folders to a user or group.

### DFSRPermissions
The script can create a delegation to a DFSR replication group for a user or group.

## Configuration Definition
This script reads its configuration from a JSON [configuration](#parameters) file with the following schema. See [ozo-ad-manage-delegations-EXAMPLE.json](https://github.com/onezeroone-dev/ozo-ad-manage-delegations/blob/main/ozo-ad-manage-delegations-EXAMPLE.json) for an example.

```json
{
    "OUDelegations":[
        {
            "Description":"",
            "Identities":[""],
            "Permissions":[""],
            "OUs":[
                ""
            ]
        }
    ],
    "GPOPermissions":[
        {
            "Description":"",
            "GPONames":[""],
            "GroupNames":[""],
            "Permissions":[""]
        }
    ],
    "DFSNRootPermissions":[
        {
            "Description":"",
            "DFSNRoots":[""],
            "Identities":[""]
        }
    ],
    "DFSNFolderPermissions":[
        {
            "Description":"",
            "DFSNFolders":[""],
            "Identities":[""]
        }
    ],
    "DFSRPermissions":[
        {
            "Description":"",
            "DFSRGroups":[""],
            "Identities":[""]
        }
    ],
    "DisabledComputersOUDN":"",
    "DisabledGroupsOUDN":"",
    "DisabledUsersOUDN":""
}

```

### Main Configuration
| Key | Value |Required|
|-----|-------|--------|
|`OUDelegations`|A list of delegations to create. See _OUDelegations Configuration_, below.|FALSE|
|`GPOPermissions`|A list of GPO permissions to apply. See _GPO Permissions Configuration_, below.|FALSE|
|`DFSNRootPermissions`|A list of DFSN roots to delegate. See _DFSN Root Permissions Configuration_, below.|FALSE|
|`DFSNFolderPermissions`|A list of DFSN folders to delegate. See _DFSN Folder Permissions Configuration_, below.|FALSE|
|`DFSRPermissions`|A list of DFSR permissions to apply. See _DFSR Permissions Configuration_, below.|FALSE|
|`DisabledComputersOUDN`|The distinguished name of the disabled computers OU.|FALSE|
|`DisabledGroupsOUDN`|The distinguished name of the disabled groups OU.|FALSE|
|`DisabledUsersOUDN`|The distinguished name of the disabled users OU.|FALSE|

### OUDelegations Configuration
|Key|Value|Required|
|---|-----|--------|
|`Description`|A brief description of the delegation.|TRUE|
|`Identities`|A list of users and groups to whom the delegation will be applied.|TRUE|
|`Permissions`|The permissions that will be applied for the user to the OU. See _Permissions_ (below) for valid values.|TRUE|
|`OUs`|A list of OU distinguished names where the delegation will be applied.|TRUE|

### GPO Permissions Configuration
The script supports applying GPO permissions only to _groups_!

|Key|Value|Required|
|---|-----|--------|
|`Description`|A brief description of the GPO permission.|TRUE|
|`GPONames`|A list of GPO names where the permissions will be applied.|TRUE|
|`GroupNames`|A list of AD groups to whom the permissions will be applied.|TRUE|
|`Permissions`|A list of the permissions to apply. Valid permissions are `GpoRead`, `GpoApply`, `GpoEdit`, and `GpoEditDeleteModifySecurity`.|TRUE|

### DFSN Root Permissions Configuration

|Key|Value|Required|
|---|-----|--------|
|`Description`|A brief description of the DFSN permission.|TRUE|
|`DFSNRoots`|A list of the folders where the permissions will be applied.|TRUE|
|`Identities`|A list of AD users and group to whom the permissions will be applied.|TRUE|

### DFSN Folder Permissions Configuration

|Key|Value|Required|
|---|-----|--------|
|`Description`|A brief description of the DFSN permission.|TRUE|
|`DFSNFolders`|A list of the folders where the permissions will be applied.|TRUE|
|`Identities`|A list of AD users and group to whom the permissions will be applied.|TRUE|

### DFSR Permissions Configuration

|Key|Value|Required|
|---|-----|--------|
|`Description`|A brief description of the DFSR permission.|TRUE|
|`DFSRGroups`|A list of the DFS replication groups where the permissions will be applied.|TRUE|
|`Identities`|A list of AD users and group to whom the permissions will be applied.|TRUE|

## Examples
```powershell
ozo-ad-manage-delegations -Configuration (Join-Path -Path $Env:USERPROFILE -ChildPath "Downloads\ozo-ad-manage-delegations.json")
```

## Logging
Messages as written to the Windows Event Viewer [_One Zero One_](https://github.com/onezeroone-dev/OZO-Windows-Event-Log-Provider-Setup/blob/main/README.md) provider when available. Otherwise, messages are written to the _Microsoft-Windows-PowerShell_ provider with event ID 4100.

## Licensing
This script is licensed under the [GNU General Public License (GPL) version 2.0](LICENSE).

## Notes
Run this script as a user with rights to create AD delegations (likely a Domain Admin) from within a writable directory.

## Relevant Links
* [DSACL docs](https://github.com/SimonWahlin/DSACL/blob/master/docs)
* [Set-GPPermission](https://learn.microsoft.com/en-us/powershell/module/grouppolicy/set-gppermission?view=windowsserver2022-ps)
* [Grant-DfsnAccess](https://learn.microsoft.com/en-us/powershell/module/dfsn/grant-dfsnaccess?view=windowsserver2022-ps)
* [Grant-DfsrDelegation](https://learn.microsoft.com/en-us/powershell/module/dfsr/grant-dfsrdelegation?view=windowsserver2022-ps)

## Acknowledgements
Special thanks to my employer, [Sonic Healthcare USA](https://sonichealthcareusa.com), who supports the growth of my PowerShell skillset and enables me to contribute portions of my work product to the PowerShell community.
