Function NewTeamsFromCSV {
  <#
    .SYNOPSIS
    Create MS Teams from CSV file.

    .DESCRIPTION
    Using this script, you can create Mirosoft Teams by importing details from csv file and this function use the MicrosoftTeams PowerShell Module.

    .EXAMPLE
    NewTeamsFromCSV -Path "C:\temp\TeamsDetails.CSV"

    Create Teams from CSV file.

    Sample CSV file format

    ```
    MailNickName,DisplayName,Description,TeamType,Owners,Members,ChannelNames,ChannelDescriptions
    TestTeam,Test Team,desc,EDU_Class,admin@Tenantname.com;helper@tenant.com,alexw@Tenantname.com,TestChannel;TestChanne2,ChannelDesc;ChannelDesc2
    TestTeam2,Test Team 2,desc1,EDU_Class,alexw@Tenantname.com,admin@Tenantname.com,TestChanne3,ChannelDesc1
    TestTeam3,Test Team 3,desc2,EDU_Class,alland@Tenantname.com,admin@Tenantname.com;AlexW@Tenantname.com;LeeG@Tenantname.com,TestChanne2,ChannelDesc3
    ```

    You need to use ";" for multiple Owners, Members, ChannelNames and ChannelDescriptions in the CSV file
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    # Provide the CSV file path of the new Teams details.
    $Path
  )

  function CreateChannel {
    param (
      $ChannelName, $ChannelDesc, $GroupId
    )
    Process {
      try {
        $teamchannels = $ChannelName -split ";"
        $descchannels = $ChannelDesc -split ";"
        if ($teamchannels) {
          for ($i = 0; $i -lt $teamchannels.count; $i++) {
            if ($i -lt $descchannels.count) {
              New-TeamChannel -GroupId $GroupId -DisplayName $teamchannels[$i] -Description $descchannels[$i]
            }
            else {
              New-TeamChannel -GroupId $GroupId -DisplayName $teamchannels[$i]
            }
          }
        }
      }
      Catch {
      }
    }
  }

  function Add-Users {
    param(
      $Users, $GroupId, $CurrentUsername, $Role
    )
    Process {

      try {
        $teamusers = $Users -split ";"
        if ($teamusers) {
          for ($j = 0; $j -lt $teamusers.count; $j++) {
            if ($teamusers[$j] -ne $CurrentUsername) {
              Add-TeamUser -GroupId $GroupId -User $teamusers[$j] -Role $Role
            }
          }
        }
      }
      Catch {
      }
    }
  }

  function CreateNewTeam {
    param (
      $ImportPath
    )
    Process {
      Import-Module MicrosoftTeams
      $cred = Get-Credential
      $username = $cred.UserName
      Connect-MicrosoftTeams -Credential $cred
      $teams = Import-Csv -Path $ImportPath
      foreach ($team in $teams) {
        $getteam = get-team | where-object { $_.displayname -eq $team.DisplayName }
        If ($null -eq $getteam) {
          Write-Host "Start creating the team: " $team.DisplayName
          $group = New-Team -MailNickName $team.MailNickName -DisplayName $team.DisplayName -Description $team.Description -Template $team.TeamType
          Write-Host "Creating channels..."
          CreateChannel -ChannelName $team.ChannelNames -ChannelDesc $team.ChannelDescriptions -GroupId $group.GroupId
          Write-Host "Adding team members..."
          Add-Users -Users $team.Members -GroupId $group.GroupId -CurrentUsername $username -Role Member
          Write-Host "Adding team owners..."
          Add-Users -Users $team.Owners -GroupId $group.GroupId -CurrentUsername $username -Role Owner
          Write-Host "Completed creating the team: " $team.DisplayName
          $team = $null
        }
        # Read-Host -Prompt "Press Enter to continue"
      }
    }
  }

  CreateNewTeam -ImportPath $Path
}

Function ExportArchivedTeams {
  <#
    .SYNOPSIS
    Export Archived Teams into CSV file.

    .DESCRIPTION
    Using this script, you can Export archived Mirosoft Teams into csv file and this function use the SharePointPnPPowerShellOnline Module.

    .EXAMPLE
    ExportArchivedTeams -ExportPath "C:\temp\ArchivedTeamsDetails.CSV"

    Export Archived Teams into CSV file.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    # Provide the CSV file path to export Archived Teams details.
    $ExportPath
  )
  process {
    Connect-PnPOnline -Scopes "Group.Read.All", "Group.ReadWrite.All"
    $accesstoken = Get-PnPAccessToken
    $group = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri  "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/any(c:c+eq+`'Team`')" -Method Get
    $TeamsList = @()

    $i = 1
    do {
      foreach ($value in $group.value) {
        Write-Progress -Activity "Get All Groups" -status "Found Group $i" -percentComplete ($i / $group.value.count * 100)

        $id = $value.id
        Try {
          $team = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri  https://graph.microsoft.com/beta/teams/$id -Method Get
          if ($null -ne $team -and $team.isArchived -eq $true) {
            $Teams = "" | Select-Object "TeamsName", "TeamType"

            $Teams.TeamsName = $value.displayname
            $Teams.TeamType = $value.visibility

            $TeamsList += $Teams
            $Teams = $null
          }
        }
        Catch {
          $team = $null
        }
        $i++
      }

      if ($null -eq $group.'@odata.nextLink') {
        break
      }
      else {
        $group = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri $group.'@odata.nextLink' -Method Get
      }
    }while ($true);
    $TeamsList
    $TeamsList | Export-csv $ExportPath -NoTypeInformation
  }
}

Function ExportTeamsList {
  <#
    .SYNOPSIS
    Export Microsoft Teams into CSV file.

    .DESCRIPTION
    Using this script, you can Export Mirosoft Teams into csv file and this function use the SharePointPnPPowerShellOnline Module.

    .EXAMPLE
    ExportTeamsList -ExportPath "C:\temp\TeamsList.CSV"

    Export Microsoft Teams into CSV file.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    # Provide the CSV file path to export Teams details.
    $ExportPath
  )

  process {
    Connect-PnPOnline -Scopes "Group.Read.All", "User.ReadBasic.All"
    $accesstoken = Get-PnPAccessToken
    $MTeams = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri  "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/any(c:c+eq+`'Team`')" -Method Get
    $TeamsList = @()
    $i = 1
    do {
      foreach ($value in $MTeams.value) {

        Write-Progress -Activity "Get All Teams" -status "Found Teams $i"

        $id = $value.id
        Try {
          $team = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri  https://graph.microsoft.com/beta/Groups/$id/channels -Method Get

        }
        Catch {

        }

        $Owner = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri  https://graph.microsoft.com/v1.0/Groups/$id/owners -Method Get
        $Members = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri  https://graph.microsoft.com/v1.0/Groups/$id/Members -Method Get
        $Teams = "" | Select-Object "TeamsName", "TeamType", "Channelcount", "ChannelName", "Owners", "MembersCount", "Members"
        $Teams.TeamsName = $value.displayname
        $Teams.TeamType = $value.visibility
        $Teams.ChannelCount = $team.value.id.count
        $Teams.ChannelName = $team.value.displayName -join ";"
        $Teams.Owners = $Owner.value.userPrincipalName -join ";"
        $Teams.MembersCount = $Members.value.userPrincipalName.count
        $Teams.Members = $Members.value.userPrincipalName -join ";"
        $TeamsList += $Teams
        $teamaccesstype = $null
        $errorMessage = $null
        $Teams = $null
        $team = $null
        $i++
      }
      if ($null -eq $MTeams.'@odata.nextLink') {
        break
      }
      else {
        $MTeams = Invoke-RestMethod -Headers @{Authorization = "Bearer $accesstoken" } -Uri $MTeams.'@odata.nextLink' -Method Get
      }
    }while ($true);
    $TeamsList
    $TeamsList | Export-csv $ExportPath -NoTypeInformation
  }
}

Function ApplyTeamsPolicyFromCSV {
  <#
    .SYNOPSIS
    Apply teams policy to user accounts from CSV file.

    .DESCRIPTION
    Using this script, you can Apply teams policy to user accounts and this function use the SkypeOnlineConnector Module.

    .EXAMPLE
    ApplyTeamsPolicyFromCSV -Path "C:\temp\TeamsPolicy.CSV"

    Apply teams policy to user account from CSV file.

    Sample CSV file format
    Ex:

    UserPricipalName,TeamsCallingPolicy,TeamsMeetingPolicy,TeamsMessagingPolicy,TeamsUpgradePolicy
    admin@tenantname.com,Tag:AllowCalling,Tag:AllOn,Tag:Default,Tag:SfBWithTeamsCollabWithNotify
    alexw@tenantname.com,Tag:DisallowCalling,Global,Tag:EduFaculty,Global
    alland@tenantname.com,Global,Global,Global,Global

    Note: CSV Header: UserPricipalName,TeamsCallingPolicy,TeamsMeetingPolicy,TeamsMessagingPolicy,TeamsUpgradePolicy
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $True)]
    [string]$Path
    # Provide the CSV file path of the teams policy details.
  )

  #connect Csonline
  $credential = Get-Credential
  $session = New-CsOnlineSession -Credential $credential
  Import-PSSession $session

  # Importing policy from csv file
  $Users = Import-Csv $Path
  foreach ($User in $Users) {

    $UserPrincipalName = $User.UserPricipalName
    $CallingPolicy = $User.TeamsCallingPolicy
    $MeetingPolicy = $User.TeamsMeetingPolicy
    $MessagingPolicy = $User.TeamsMessagingPolicy
    $UpgradePolicy = $User.TeamsUpgradePolicy

    Get-CsOnlineUser | ForEach-Object {
      $OnlineUser = $_
      #Apply teams policy
      if ($userprincipalname -eq $OnlineUser.userprincipalname) {

        if ($CallingPolicy -eq 'Global') {
          Grant-CsTeamsCallingPolicy   -Identity $UserPrincipalName -PolicyName $null
        }
        else {
          Grant-CsTeamsCallingPolicy   -Identity $UserPrincipalName -PolicyName $CallingPolicy
        }
        if ($MeetingPolicy -eq 'Global') {
          Grant-CsTeamsMeetingPolicy   -Identity $UserPrincipalName -PolicyName $null
        }
        else {
          Grant-CsTeamsMeetingPolicy   -Identity $UserPrincipalName -PolicyName $MeetingPolicy
        }
        if ($MessagingPolicy -eq 'Global') {
          Grant-CsTeamsMessagingPolicy -Identity $UserPrincipalName -PolicyName $null
        }
        else {
          Grant-CsTeamsMessagingPolicy -Identity $UserPrincipalName -PolicyName $MessagingPolicy
        }
        if ($UpgradePolicy -eq 'Global') {

          Grant-CsTeamsUpgradePolicy   -Identity $UserPrincipalName -PolicyName $null
        }
        else {
          Grant-CsTeamsUpgradePolicy   -Identity $UserPrincipalName -PolicyName $UpgradePolicy
        }
      }
      Else {
        Write-Host  $UserPrincipalName "user is not available in the Azure Active directory"
      }
    }
  }
  #Creating some delay
  Start-Sleep -Seconds 90

  #Checking given policy is applied or not.
  foreach ($User in $Users) {
    $UserPrincipalName = $User.UserPricipalName
    Get-CsOnlineUser | ForEach-Object {
      $OnlineUser = $_
      if ($userprincipalname -eq $OnlineUser.userprincipalname) {
        New-Object -TypeName PSObject -Property @{
          UserPrincipalname    = $OnlineUser.userprincipalname
          TeamsCallingPolicy   = $OnlineUser.TeamsCallingPolicy
          TeamsMeetingPolicy   = $OnlineUser.TeamsMeetingPolicy
          TeamsMessagingPolicy = $OnlineUser.TeamsMessagingPolicy
          TeamsUpgradePolicy   = $OnlineUser.TeamsUpgradePolicy
        }
      }
    } | Select-Object userprincipalname, TeamsCallingPolicy, TeamsMeetingPolicy, TeamsMessagingPolicy, TeamsUpgradePolicy
  }
}

Function ExportTeamsPolicy {
  <#
    .SYNOPSIS
    Export Microsoft Teams policy from the user accounts into CSV file.

    .DESCRIPTION
    Using this script, you can Export Mirosoft Teams policy into csv file and this function use the SkypeOnlineConnector Module.

    .EXAMPLE
    ExportTeamsPolicy -Path "C:\temp\TeamsPolicyDetails.CSV"

    Export Microsoft Teams policy from the user accounts into CSV file.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    # Provide the CSV file path to export Teams policy details.
    $Path
  )

  #Connect Csonline
  $credential = Get-Credential
  $session = New-CsOnlineSession -Credential $credential
  Import-PSSession $session

  #Get CsOnline User
  $GetTeams = Get-CsOnlineUser | ForEach-Object {
    $OnlineUser = $_
    New-Object -TypeName PSObject -Property @{
      UserPrincipalname    = $OnlineUser.userprincipalname
      TeamsCallingPolicy   = If ($null -eq $OnlineUser.TeamsCallingPolicy) { "Global" } else { $OnlineUser.TeamsCallingPolicy }
      TeamsMeetingPolicy   = If ($null -eq $OnlineUser.TeamsMeetingPolicy) { "Global" } else { $OnlineUser.TeamsMeetingPolicy }
      TeamsMessagingPolicy = If ($null -eq $OnlineUser.TeamsMessagingPolicy) { "Global" } else { $OnlineUser.TeamsMessagingPolicy }
      TeamsUpgradePolicy   = If ($null -eq $OnlineUser.TeamsUpgradePolicy) { "Global" } else { $OnlineUser.TeamsUpgradePolicy }

    }
  } | Select-Object userprincipalname, TeamsCallingPolicy, TeamsMeetingPolicy, TeamsMessagingPolicy, TeamsUpgradePolicy

  $GetTeamsPolicy
  $GetTeamsPolicy | Export-Csv -Path $Path -NoTypeInformation
}
