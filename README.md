## XEws - Powershell module for working with Exchange Web Services

Powershell module leveraging EWS managed api. Allows accessing to mailboxes using EWS calls.

### Module installation

Invoke-Expression -Command (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/IvanFranjic/XEws/master/InstallXEwsModule.ps1")


### Basic usage

Import session:

```
Import-EwsSession -EmailAddress "user@domain.com" -Password (ConvertTo-SecureString -AsPlainText -Force -String "pass") -EwsEndpoint "https://mail.domain.com/ews/exchange.asmx"
```

Start using it:

```
Get-EwsFolder
```

```
Get-EwsItem
```