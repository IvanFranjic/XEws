# create autoupdate script for xews
[CmdletBinding()]
param
()

function Get-EwsOnlineVersion
{
    [CmdletBinding()]
    param
    (
        $ModuleHomePath
    )

    try
    {
        $moduleTempManifestPath = [String]::Format('{0}\{1}', $ModuleHomePath, 'XEws-OnlineVersion.psd1');
        $webClient = New-Object System.Net.WebClient;        
        $webClient.DownloadFile('https://raw.githubusercontent.com/IvanFranjic/XEws/master/bin/Debug/XEws.psd1', $moduleTempManifestPath);
        Get-EwsModuleVersion -ModuleData (Get-Content $moduleTempManifestPath | Select-String -Pattern 'ModuleVersion' -CaseSensitive);
    }
    catch
    {
        Write-EwsUpdateLog -Message "Error encountered: $($_.Exception.Message)" -CallingFunction "Get-EwsOnlineVersion";
    }
    
    [System.IO.File]::Delete($moduleTempManifestPath);
}

function Get-EwsLocalVersion
{
    [CmdletBinding()]
    param
    (
        $ModuleHomePath
    )

    $moduleManifestPath = [String]::Format('{0}\{1}', $ModuleHomePath, 'XEws.psd1');
    Get-EwsModuleVersion -ModuleData (Get-Content $moduleManifestPath | Select-String -Pattern 'ModuleVersion' -CaseSensitive);
}

function Get-EwsModuleVersion
{
    param
    (
        [string]
        $ModuleData
    )

    if (-not ([string]::IsNullOrEmpty($ModuleData)))
    {
        $stringVersion = ($ModuleData.Split('=')[1].Trim().Replace("'", ''));
        return ([double]$stringVersion)
    }
    return $null;
}

function Write-EwsUpdateLog
{
    param
    (
        [string]
        $Message,

        [string]
        $CallingFunction
    )

    $LogPath = $Global:ewsLogPath;

    if (Test-Path -Path $LogPath)
    {
        $logFileSize = Get-Item -Path $LogPath;
        if ($logFileSize.Length -gt 1500)
        {
            Remove-Item -Path $LogPath | Out-Null;
        }
    }

    $messageToWrite = [string]::Format(
        "[{0} - {1}] - $Message",
        (Get-Date).ToString("hh:mm:ss dd-MM-yyyy"),
        $CallingFunction
    );

    $streamWriter = New-Object System.IO.StreamWriter($LogPath, $true);
    $streamWriter.WriteLine($messageToWrite);
    $streamWriter.Close();
    $streamWriter.Dispose();
}


# --- Main --- #

$xewsModuleHomePath = [String]::Format('{0}\XEws', $env:PSModulePath.Split(';') -match $env:USERNAME);
$Global:ewsLogPath = ([System.IO.Path]::Combine($xewsModuleHomePath, "EwsUpdateLog.txt"));

try
{
    Write-EwsUpdateLog -Message "Comparing local and online version" -CallingFunction "Script Root";

    $localVersion = Get-EwsLocalVersion -ModuleHomePath $xewsModuleHomePath -ErrorAction Stop;
    $onlineVersion = Get-EwsOnlineVersion -ModuleHomePath $xewsModuleHomePath -ErrorAction Stop;

    if ($localVersion -lt $onlineVersion)
    {
        Write-EwsUpdateLog -Message "Detected new module build online. Local version: $localVersion, Remote version: $onlineVersion" -CallingFunction "Script Root";
        Invoke-Expression -Command (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/IvanFranjic/XEws/master/InstallXEwsModule.ps1');
    }
    else
    {
        Write-EwsUpdateLog -Message 'Latest version already installed.' -CallingFunction "Script Root";
    }
}
catch
{
    Write-EwsUpdateLog -Message "Could not download newer module version. Error: $($_.Exception.Message)." -CallingFunction "Script Root";
}

Add-Type -Path ([System.IO.Path]::Combine($xewsModuleHomePath, "Microsoft.Exchange.WebServices.dll"));