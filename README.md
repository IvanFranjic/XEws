## XEws - Powershell module for working with Exchange Web Services

Powershell module leveraging EWS managed api for working with mailboxes.

### Module installation

```
Invoke-Expression -Command ((Invoke-WebRequest "https://raw.githubusercontent.com/IvanFranjic/XEws/master/InstallXEwsModule.ps1").Content)
```

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