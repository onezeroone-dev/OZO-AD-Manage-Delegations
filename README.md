# OZO AD Manage Delegations Installation and Usage
## Description
Creates AD delegations based on a configuration file. This script can create OU Delegations, apply permissions to GPOs, grant access to DFSN roots, grant access to DFSN folders, and create delegations to DFSR replication groups. The JSON configuration file and resulting report can be useful as documentation and as evidence for a security audit.

## Prerequisites
This script requires the _ActiveDirectory_, _DFSN_, _DFSR_, _DSACL_, _GroupPolicy_, _ImportExcel_, _OZO_, _OZOAD_, _OZOFiles_,and _OZOLogger_ PowerShell modules. The ActiveDirectory, DFSN, DFSR, and GroupPolicy modules are included with the [Remote Server Administration Tools](https://learn.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools) installation. The remaining modules are published to [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Module DSACL,ImportExcel,OZO,OZOAD,OZOFiles,OZOLogger
```

## Installation
This script is published to Microsoft's [PowerShell Gallery](https://learn.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-5.1). Ensure your system is configured for this repository then execute the following in an _Administrator_ PowerShell:

```powershell
Install-Script ozo-ad-manage-delegations
```

## Usage
```
ozo-ad-manage-delegations
    -Configuration <String>
    [-OutDir <String>]
```

## Parameters
|Parameter|Description|
|---------|-----------|
|`Configuration`|Path to the JSON configuration file. Defaults to `ozo-ad-manage-delegations.json` in the same directory as the script. Please see _Configuration Definition_ (below) for more information.|
|`OutDir`|Directory for the Excel tesults report. Defaults to the current directory.|

## Capabilities
### OU Delegations
The following permissions are implemented in the script. See [_Configuration Definition_](#configuration-definition) below for information on how to use these permissions.

|Permission|Description|
|----------|-----------|
|`CreateChildComputers`|Delegates create child computer objects.|
|`CreateChildContacts`|Delegates create child contact objects.|
|`CreateChildGroups`|Delegates create child group objects.|
|`CreateChildUsers`|Delegates create child user objects.|
|`CreateOUs`|Delegates create OU.|
|`DeleteChildComputers`|Delegates delete child computer objects.|
|`DeleteChildContacts`|Delegates delete child contact objects.|
|`DeleteChildGroups`|Delegates delete child group objects.|
|`DeleteChildUser`|Delegates delete child user objects.|
|`DeleteOUs`|Delegates delete OU.|
|`DomainJoinComputer`|Delegates create computer objects, Write Name, and Write name.|
|`EnableDisableComputers`|Delegates enable and disable computer objects.|
|`EnableDisableUsers`|Delegates enable and disable user objects|
|`FullControlComputers`|Delegates full control to computer objects.|
|`FullControlContacts`|Delegates full control to contact objects.|
|`FullControlGroups`|Delegates full control to group objects.|
|`FullControlUsers`|Delegates full control to user objects.|
|`FullControlOUs`|Delegates full control to organizational unit objects.|
|`LinkGPO`|Delegates permission to link GPOs.|
|`ModifyGroupMembership`|Delegates modify group membership.|
|`ReadBitLockerRecovery`|Delegates read to the BitLocker Recovery information.|
|`ResetUserPasswords`|Delegates reset password.|

Note: To allow an identity to move computers from one OU to another, delegate `DeleteChildComputers` on the source OU and `CreateChildComputers` on the target OU; and likewise for contacts, groups, and users.

### GPO Permissions
The script can apply `GpoRead`, `GpoApply`, `GpoEdit`, and `GpoEditDeleteModifySecurity` to _groups_.

### DFSN Root Permissions
The script can grant access to a DFSN **root** to a user or group.

### DFSN Folder Permissions
The script can grant access to DFSN folders to a user or group.

### DFSR Permissions
The script can create a delegation to a DFSR replication group for a user or group.

## JSON Configuration Definition
This script leverages the [One Zero One Unified AD JSON Schema](https://onezeroone.dev/ozo-unified-ad-json-schema/). The elements of the schema used by this script are as follows. Please also see [ozo-ad-manage-delegations-EXAMPLE.json](https://github.com/onezeroone-dev/ozo-ad-manage-delegations/blob/main/ozo-ad-manage-delegations-EXAMPLE.json).

## Examples
```powershell
ozo-ad-manage-delegations -Configuration (Join-Path -Path $Env:USERPROFILE -ChildPath "Downloads\ozo-ad-manage-delegations.json")
```

## Logging
Messages as written to the Windows Event Viewer [_One Zero One_](https://github.com/onezeroone-dev/OZO-Windows-Event-Log-Provider-Setup/blob/main/README.md) provider when available. Otherwise, messages are written to the _Microsoft-Windows-PowerShell_ provider with event ID 4100.

## Licensing
This script is licensed under the [GNU General Public License (GPL) version 2.0](LICENSE).

## Notes
Run this script as a user with rights to create AD delegations (e.g., a Domain Admin).

## Relevant Links
* [DSACL docs](https://github.com/SimonWahlin/DSACL/blob/master/docs)
* [Set-GPPermission](https://learn.microsoft.com/en-us/powershell/module/grouppolicy/set-gppermission?view=windowsserver2022-ps)
* [Grant-DfsnAccess](https://learn.microsoft.com/en-us/powershell/module/dfsn/grant-dfsnaccess?view=windowsserver2022-ps)
* [Grant-DfsrDelegation](https://learn.microsoft.com/en-us/powershell/module/dfsr/grant-dfsrdelegation?view=windowsserver2022-ps)

## Acknowledgements
Special thanks to my employer, [Sonic Healthcare USA](https://sonichealthcareusa.com), who supports the growth of my PowerShell skillset and enables me to contribute portions of my work product to the PowerShell community.
