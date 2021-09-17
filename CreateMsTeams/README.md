# Bulk create teams and channels

## Prerequisite

```powershell
Set-ExecutionPolicy RemoteSigned
Install-Module -Name MicrosoftTeams
```

## CSV template

- MailNickName must not contain whitespace
- TeamType of A1 can only be `EDU_Class` or `EDU_PLC`
- Owners, Members, ChannelNames, ChannelDescriptions can be an array whose elements are separated using a semicolon `;`

## Credits

[Marius Pretorius](https://techcommunity.microsoft.com/t5/microsoft-teams-for-education/teams-admin-creating-bulk-class-teams-using-powershell-and-a-csv/m-p/1931633)

[New-Team](https://docs.microsoft.com/en-us/powershell/module/teams/new-team)

[New-TeamChannel](https://docs.microsoft.com/en-us/powershell/module/teams/new-teamchannel)
