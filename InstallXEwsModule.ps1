[CmdletBinding()]
param()

$moduleFiles = @("Microsoft.Exchange.WebServices.dll", "Microsoft.Exchange.WebServices.xml", "XEws.dll", "XEws.format.ps1xml", "XEws.psd1", "en-US/XEws.dll-Help.xml");
$githubEwsEndpoint = "https://github.com/IvanFranjic/XEws/tree/master/bin/Debug";
$moduleHomeFolder = [String]::Format("{0}\XEws", $env:PSModulePath.Split(";") -match $env:USERNAME);

Write-Verbose -Message "Checking if module home folder exist...";

if (!(Test-Path $moduleHomeFolder))
{
    try
    {
        $null = New-Item -ItemType "Directory" -Path ($env:PSModulePath.Split(";")[0]) -Name "XEws" -Force -ErrorAction Stop;
        $null = New-Item -ItemType "Directory" -Path $moduleHomeFolder -Name "en-US" -Force;
    }
    catch
    {
        throw "Error encountered: $($_.Exception.Message)";
    }
}

Write-Verbose -Message "Initializing web client...";

$webClient = New-Object System.Net.WebClient;

foreach ($item in $moduleFiles)
{
    $moduleFileLocation = [String]::Format("{0}/{1}", $githubEwsEndpoint, $item);
        
    if ($item.Contains("/"))
    {
        $item = $item.Replace("/", "\");
    }

    Write-Verbose -Message "Downloading file $item";

    $fileLocation = [String]::Format("{0}\{1}", $moduleHomeFolder, $item);
    $fileData = $webClient.DownloadData($moduleFileLocation);       
    Set-Content -Value $fileData -Path $fileLocation -Encoding Byte -Force;
}
