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
        $webClient.DownloadFile('https://raw.githubusercontent.com/IvanFranjic/XEws/master/bin/Debug/XEws.psd1', $moduleTempManifestPath)
        Get-EwsModuleVersion -ModuleData (Get-Content $moduleTempManifestPath | Select-String -Pattern 'ModuleVersion' -CaseSensitive);
    }
    catch
    {
        throw "[Get-EwsOnlineVersion] -  Error occured. $($_.Exception.Message)."
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

# --- Main --- #

$xewsModuleHomePath = [String]::Format('{0}\XEws', $env:PSModulePath.Split(';') -match $env:USERNAME);

try
{
    Write-Warning -Message 'Checking if newer version of module exist...';

    $localVersion = Get-EwsLocalVersion -ModuleHomePath $xewsModuleHomePath -ErrorAction Stop;
    $onlineVersion = Get-EwsOnlineVersion -ModuleHomePath $xewsModuleHomePath -ErrorAction Stop;

    if ($localVersion -lt $onlineVersion)
    {
        Write-Warning -Message 'Detected new module build online. Downloading, please wait...';
        Invoke-Expression -Command (Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/IvanFranjic/XEws/master/InstallXEwsModule.ps1');
    }
    else
    {
        Write-Warning -Message 'Latest version already installed...'
    }
}
catch
{
    Write-Warning -Message "Could not download newer module version. Error: $($_.Exception.Message)."
}
