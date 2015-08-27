## XEws - Powershell module for working with Exchange Web Services

Powershell module leveraging EWS managed api for working with mailboxes.

### Module installation

Download InstallXEwsModule.ps1 script and run it.


### Basic usage

Import session:

```
Import-XEwsSession -UserName "user@domain.com" -Password "P@ssw0rd" -EwsUri "https://mail.domain.com/ews/exchange.asmx"
```

Start using it:

```
Get-XEwsFolder
```

```
Get-XEwsItem
```