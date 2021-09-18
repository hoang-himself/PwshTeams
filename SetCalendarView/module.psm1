Function SetCalendarFromCSV {
  <#
    .SYNOPSIS
    Set calendar configuration of all user from a CSV.

    .DESCRIPTION
    Using this script, you can edit the calendar view of all users defined in a CSV.

    .EXAMPLE
    SetCalendarFromCSV -Path "C:\temp\Users.CSV"

    Create Teams from CSV file.

    Sample CSV file format

    ```
    Username
    admin@Tenantname.com
    alexw@Tenantname.com
    alland@Tenantname.com
    ```
    or
    ```
    Username
    admin
    alexw
    alland
    ```
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]
    # Provide the CSV file path of the new Teams details.
    $Path
  )
  Import-Module ExchangeOnlineManagement
  $cred = Get-Credential
  Connect-ExchangeOnline -Credential $cred

  $users = Import-Csv -Path $Path
  foreach ($user in $users) {
    Set-MailboxCalendarConfiguration -Identity $user.Username -FirstWeekOfYear FirstFullWeek `
      -ShowWeekNumbers $true -WeekStartDay "Monday" -WorkDays "Weekdays,Saturday" `
      -WorkingHoursStartTime 06:45:00 -WorkingHoursEndTime 16:55:00 `
      -DefaultReminderTime 00:10:00 -WorkingHoursTimeZone "SE Asia Standard Time"
  }
}

Function SetCalendarAllUsers {
  <#
    .SYNOPSIS
    Set calendar configuration of all user.

    .DESCRIPTION
    Using this script, you can edit the calendar view of all users.

    .EXAMPLE
    SetCalendarAllUsers
  #>
  Import-Module ExchangeOnlineManagement
  $cred = Get-Credential
  Connect-ExchangeOnline -Credential $cred

  Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox |
  Set-MailboxCalendarConfiguration -FirstWeekOfYear FirstFullWeek `
    -ShowWeekNumbers $true -WeekStartDay "Monday" -WorkDays "Weekdays,Saturday" `
    -WorkingHoursStartTime 06:45:00 -WorkingHoursEndTime 16:55:00 `
    -DefaultReminderTime 00:10:00 -WorkingHoursTimeZone "SE Asia Standard Time"
}
