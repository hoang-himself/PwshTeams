Import-Module ExchangeOnlineManagement
$cred = Get-Credential
Connect-MicrosoftTeams -Credential $cred

Get-Mailbox -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq "UserMailbox" } |
Set-MailboxCalendarConfiguration -FirstWeekOfYear FirstFullWeek `
  -ShowWeekNumbers $true -WeekStartDay Monday -WorkDays "Weekdays,Saturday" `
  -WorkingHoursStartTime 06:45:00 -WorkingHoursEndTime 16:55:00 `
  -DefaultReminderTime 00:10:00 -WorkingHoursTimeZone "SE Asia Standard Time"
