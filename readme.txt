
# Mail2Ticket
## Office / Outlook 2024
## Copy Mails to Ticket Pickup folder and change Subject to Ticket-ID


### Exchange installation
[PS] C:\Windows\system32>New-App -OrganizationApp -FileData ([System.IO.File]::ReadAllBytes("C:\OToTicket\manifest.xml"))

[PS] C:\Windows\system32>Remove-App -OrganizationApp -Identity "12345678-1234-1234-1234-123456789012"

Bestätigung
Möchten Sie diese Aktion wirklich ausführen?
Die App "12345678-1234-1234-1234-123456789012" wird aus der Organisation deinstalliert.
[J] Ja  [A] Ja, alle  [N] Nein  [K] Nein, keine  [?] Hilfe (Standard ist "J"): j

Set-App -Identity "OToTicket" -Enabled $true -DefaultStateForUser Enabled


# Alle Organisations-Apps anzeigen
Get-App -OrganizationApp

# Alle Apps für einen Benutzer anzeigen  
Get-App -Mailbox "ulewu@chaos.local"


### DB API

localhost:8080/api/tickets/suggestions?q=test

