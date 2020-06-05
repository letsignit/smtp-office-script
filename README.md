# PowerShell Script for Office LSI Connector

`smtp-office.ps1` is a PowerShell script to help in the configuration of Office 365 to work with Letsignit SMTP.
You must install at least PowerShell version 4.
See [here](https://docs.microsoft.com/en-us/powershell/scripting/install/installing-windows-powershell?view=powershell-6)

`smtp-office-us.ps1` does the same but target US Letsignit SMTP infrastructure.

`smtp-office-ca.ps1` does the same but target CAD Letsignit SMTP infrastructure.

`smtp-office-mfa.ps1` and `smtp-office-us-mfa.ps1` and `smtp-office-ca-mfa.ps1` use authentication scheme that is compatible with MFA.

## Usage 

Before using the script, it is necessary that the SMTP functionality is active on the LSI application for the chosen domain.

Before using this script, you must have global administrator rights.

Just launch `smtp-office.ps1`. 

The script will do the following operations:
 - ask for a login password
 - ask for the domain to be configured
 - ask for an email that will be used for testing the configuration
 - check that the domain is an accepted domain
 - add an inbound connector that allows the Letsignit SMPT to route emails to Office 365 to then be delivered to the receiver
 - add an outbound connector that allows routing incoming email to Letsignit SMTP
 - add a transport rule that permits filtering emails that should use the outbound connector and an exception
   for Out Of Office message.
 - add a connection filter policy that adds LSI SMTP IP in order to prevent emails from LSI from being considered as spam
 - disable rich text format that is not compatible with LSI
 - fix incorrect distribution groups by adding a ReportToOriginator option

At each step, the script will try to detect previous LSI configurations. 
If an element seems to contain LSI information, 
it will be updated.

When the script is finished, nothing is enabled, and the added items do not affect the existing one. 
To make the configuration functional, it is necessary to activate the transport rule.


## Multiple domains

If you need to configure multiple domains, just activate the domain on the LSI application and run the script for each domain.

## Note for MFA

Scripts with MFA compatible authentication use the Connect-EXOPSSession cmdlet and 
Exchange Online Remote PowerShell Module [learn more](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps).

For the sake of simplicity, the script install, if needed, and load the module. 

To do that you need to enabled running unsigned script by doing

```
set-executionpolicy remotesigned
```

After Letsignit's operation completed, you could revert this setting by doing

```
set-executionpolicy restricted
```

