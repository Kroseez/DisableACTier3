# DisableACTier3
The third part. Automate the locking of terminated employees' accounts in the active directory and Microsoft Exchange. Script PowerShell .PS1
The latest iteration of the script from the aforementioned project, the objective of which is to automate the process of disabling user credentials.
The script's functionality encompasses the identification of all mailboxes within a designated OU, 
their subsequent migration to the terminated employees database in Exchange, 
the disablement of Exchange ActiveSync and Outlook Web App for these users, and the concealment of these mailboxes from the address book (GAL). Additionally, the script is designed to transmit migration notifications to mail.

Required:
- ActiveDirectory module
- AD and Exchange administrator rights
