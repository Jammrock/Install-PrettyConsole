#Requires -RunAsAdministrator

# sets up pretty console based on Scot Hanselman's pretty console articles: https://www.hanselman.com/blog/my-ultimate-powershell-prompt-with-oh-my-posh-and-the-windows-terminal
# updated 1 Sept 2021

[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $Path = "C:\Temp"
)

## CONSTANTS ##

# repro for Caskaydia Cove Nerd Font
$repoCCNF = "ryanoasis/nerd-fonts"

# name of the preferred pretty font, CaskaydiaCove NF
$fontName = "CaskaydiaCove NF"

# the zip file where CC NF is in
$fontFile = "CascadiaCode.zip"

# list of packages that will be installed or upgrades with winget. Must be exact match, as --exact is used with all winget calls.
# PowerShell and Terminal must be the last two packages, in that order, to prevent errors.
[string[]]$Packages = "JanDeDobbeleer.OhMyPosh", "Microsoft.PowerShell", "Microsoft.WindowsTerminal"

# list of commands to add to the PowerShell profile
[string[]]$profileLines = 'Import-Module -Name Terminal-Icons',
                          'oh-my-posh --init --shell pwsh --config ~/jandedobbeleer.omp.json | Invoke-Expression'


# install font for all WT profiles, or just PowerShell
# true = PowerShell [7] only
# false = All WT profiles
[bool]$wtFontPwshOnly = $true

# Windows PowerShell might be the default WT profile
# this sets PowerShell [7+] as the default profile.
[bool]$pwshDefault = $true

# get rid of annoying copy/paste prompts
[bool]$devilGetBehindMe = $true


## FUNCTIONS ##
#region

# FUNCTION: Find-GitReleaseLatest
# PURPOSE:  Calls Github API to retrieve details about the latest release. Returns a PSCustomObject with repro, version (tag_name), and download URL.
function Find-GitReleaseLatest
{
    [CmdletBinding()]
    param(
        [string]$repo
    )

    Write-Verbose "Find-GitReleaseLatest - Begin"

    $baseApiUri = "https://api.github.com/repos/$($repo)/releases/latest"


    # get the available releases
    Write-Verbose "Find-GitReleaseLatest - Processing repro: $repo"
    Write-Verbose "Find-GitReleaseLatest - Making Github API call to: $baseApiUrl"
    try 
    {
        if ($pshost.Version.Major -le 5)
        {
            $rawReleases = Invoke-WebRequest $baseApiUri -UseBasicParsing -EA Stop
        }
        elseif ($pshost.Version.Major -ge 6)
        {
            $rawReleases = Invoke-WebRequest $baseApiUri -EA Stop
        }
        else 
        {
            return (Write-Error "Unsupported version of PowerShell...?" -EA Stop)
        }
    }
    catch 
    {
        return (Write-Error "Could not get GitHub releases. Error: $_" -EA Stop)        
    }

    Write-Verbose "Find-GitReleaseLatest - Processing results."
    try
    {
        [version]$version = ($rawReleases.Content | ConvertFrom-Json).tag_name
    }
    catch
    {
        $version = ($rawReleases.Content | ConvertFrom-Json).tag_name
    }

    Write-Verbose "Find-GitReleaseLatest - Found version: $version"

    $dlURI = ($rawReleases.Content | ConvertFrom-Json).Assets.browser_download_url

    Write-Verbose "Find-GitReleaseLatest - Found download URL: $dlURI"

    Write-Verbose "Find-GitReleaseLatest - End"

    return ([PSCustomObject]@{
        Repo    = $repo
        Version = $version
        URL     = $dlURI
    })
} #end Find-GitReleaseLatest


function Get-InstalledFonts
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    return ((New-Object System.Drawing.Text.InstalledFontCollection).Families)
}


function Install-Font
{
    [CmdletBinding()]
    param (
        [Parameter()]
        $Path
    )

    $FONTS = 0x14
    $CopyOptions = 4 + 16;
    $objShell = New-Object -ComObject Shell.Application
    $objFolder = $objShell.Namespace($FONTS)

    foreach ($font in $Path)
    {
        $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
        $objFolder.CopyHere($font.fullname,$CopyFlag)
    }
}

# FUNCTION: Get-WebFile
# PURPOSE:  
function Get-WebFile
{
    param ( 
        [string]$URI,
        [string]$savePath,
        [string]$fileName
    )

    Write-Verbose "Get-WebFile - Begin"
    Write-Verbose "Get-WebFile - Attempting to download: $dlUrl"

    # make sure we don't try to use an insecure SSL/TLS protocol when downloading files
    $secureProtocols = @() 
    $insecureProtocols = @( [System.Net.SecurityProtocolType]::SystemDefault, 
                            [System.Net.SecurityProtocolType]::Ssl3, 
                            [System.Net.SecurityProtocolType]::Tls, 
                            [System.Net.SecurityProtocolType]::Tls11) 
    foreach ($protocol in [System.Enum]::GetValues([System.Net.SecurityProtocolType])) 
    { 
        if ($insecureProtocols -notcontains $protocol) 
        { 
            $secureProtocols += $protocol 
        } 
    } 
    [System.Net.ServicePointManager]::SecurityProtocol = $secureProtocols

    try 
    {
        Invoke-WebRequest -Uri $URI -OutFile "$savePath\$fileName"
    } 
    catch 
    {
        return (Write-Error "Could not download $URI`: $($Error[0].ToString())" -EA Stop)
    }

    Write-Verbose "Get-WebFile - File saved to: $savePath\$fileName"
    Write-Verbose "Get-WebFile - End"
    return "$savePath\$fileName"
} #end Get-WebFile

function Install-WGPackage
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [string[]]
        $Package
    )

    foreach ($pack in $Package)
    {
        # is the package installed already?
        $fndPack = winget list $pack --exact | Where-Object { $_ -match $pack}
        Write-Verbose "Package: $fndPack"

        # check for updates when found
        if ($fndPack)
        {
            Write-Verbose "Looking for newer version."
            # get the installed version
            [version]$packVer = ($fndPack -replace "\s+"," ").Trim(" ").Split(" ")[-1]
            Write-Verbose "Installed: $($packVer.ToString())"

            # get newest version
            [version]$newVer = winget search $pack --exact | Where-Object { $_ -match $pack } | ForEach-Object { $_.trim(" ").split(" ")[-1] }
            Write-Verbose "Available: $($newVer.ToString())"

            if ($newVer -gt $packVer)
            {
                if ($pack -match "WindowsTerminal")
                {
                    Write-Host -Foreground Yellow "`n`n`The terminal may close once Windows Terminal is upgraded. This is expected. Please re-run the script to continue.`n`nPlease run Windows Terminal with PowerShell, not Windows PowerShell, to ensure the correct profile is used.`n`n"
                }
                Write-Verbose "Upgrading $pack"
                winget upgrade $pack --exact
            }
        }
        # install if not
        else
        {
            Write-Verbose "Installing $pack"
            winget install $pack --exact
        }
    }
}

#endregion FUNCTIONS



## MAIN ##

# make sure the path is there
$null = mkdir "$Path" -Force -EA SilentlyContinue

# requires winget (installed by default on Win11 and newer Win10)
$fndWinget = Get-Command winget -EA SilentlyContinue

if ( -NOT $fndWinget )
{
    # try to install it
    $wgRelease = Find-GitReleaseLatest 'microsoft/winget-cli'
    $wgURL = $wgRelease.URL | Where-Object { $_ -match "msixbundle" }
    $wgFile = Get-WebFile -URI $wgURL -savePath $Path -fileName "AppInstaller.msixbundle"

    # install App Installer
    try 
    {
        $null = Import-Module -Name Appx -EA Stop    
    }
    catch 
    {
        $null = Import-Module -Name Appx -UseWIndowsPowershell -EA SilentlyContinue    
    }

    Add-AppxPackage -Path $wgFile

    $fndWinget = Get-Command winget -EA SilentlyContinue

    if ( -NOT $fndWinget )
    {
        return (Write-Error "Winget is required. Please install Winget from Github or upgrade to a newer version of Windows. https://github.com/microsoft/winget-cli/releases/latest")
    }
}

# get CaskaydiaCove NF if not installed
if ($fontName -notin (Get-InstalledFonts))
{
    # get newest font
    $ccnf = Find-GitReleaseLatest -repo $repoCCNF    

    # find the correct URL
    $ccnfURL = $ccnf.URL | Where-Object {$_ -match $fontFile}

    # download
    try 
    {
        $ccnfZip = Get-WebFile -URI $ccnfURL -savePath $Path -fileName $fontFile    
    }
    catch 
    {
        Write-Error "Failed to download $fontFile. Please download and install $fontName manually, or the Nerd Font of your choice."
    }
    
    # extract
    $extractPath = "$Path\ccnf"
    Expand-Archive -Path $ccnfZip -DestinationPath $extractPath -Force

    # install fonts
    Install-Font (Get-ChildItem "$extractPath" -Filter "*.ttf" -EA SilentlyContinue)

    Start-Sleep 30
}

# install/upgrade winget packages
Install-WGPackage $Packages -Verbose

# install terminal icons
Install-Module -Name Terminal-Icons -Repository PSGallery -Force

# edit profile with default Terminal Icons and posh profile
$profileLines | Out-File $PROFILE -Force -NoClobber -Append


# edit the WT settings file so it uses new font
$wtPackNames = "Microsoft.WindowsTerminal", "Microsoft.WindowsTerminalPreview"

:wt foreach ($wt in $wtPackNames) 
{
    # continue if installed
    # at a minimum, the normal WT will be installed but Preview may not
    $appxPack = Get-AppxPackage -Name $wt -EA SilentlyContinue

    #  command may fail when using PowerShell [Core], so try alternate method
    if (-NOT $appxPack)
    {
        $null = Import-Module -Name Appx -UseWIndowsPowershell -EA SilentlyContinue

        $appxPack = Get-AppxPackage -Name $wt -EA SilentlyContinue

        if (-NOT $appxPack)
        {
            Write-Warning "Failed to find the Windows Terminal path. Please manually set the font in WT to '$fontName'."
            break wt
        }
    }

    # assume WT is installed at this point
    $wtAppData = "$ENV:LOCALAPPDATA\Packages\$($appxPack.PackageFamilyName)\LocalState"

    # export settings.json
    # clean up comment lines to prevent issues with older JSON parsers (looking at you Windows PowerShell)
    $wtJSON =  Get-Content "$wtAppData\settings.json" | Where-Object { $_ -notmatch "^.*//.*$" -and $_ -ne "" -and $_ -ne $null} | ConvertFrom-Json

    if ($wtFontPwshOnly)
    {
        # change the font for PowerShell
        if ($null -ne $wtJSON.profiles.list.font.face)
        {
            $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" } | ForEach-Object { $_.Font.Face = $fontName }
        }
        else 
        {
            $pwshProfile = $wtJSON.profiles.list | Where-Object { $_.Name -eq "PowerShell" }
            $pwshProfile | Add-Member -NotePropertyName font -NotePropertyValue ([PSCustomObject]@{face="$fontName"})
        }
            
        
    }
    else 
    {
        # change the font for PowerShell
        $wtJSON.profiles.list.Font | ForEach-Object { $_.Face = $fontName }
    }

    # set PowerShell as default
    if ($pwshDefault)
    {
        $pwshGUID = $wtJSON.profiles.list | Where-Object Name -eq "PowerShell" | ForEach-Object { $_.guid }

        if ($pwshGUID)
        {
            $defaultGUID = $wtJSON.defaultProfile

            if ($defaultGUID -ne $pwshGUID)
            {
                $wtJSON.defaultProfile = $pwshGUID
            }
        }
    }

    # get rid of annoying copy prommpts
    if ($devilGetBehindMe)
    {
        $evilPasteSettings = 'largePasteWarning', 'multiLinePasteWarning'

        foreach ($imp in $evilPasteSettings)
        {
            if ($null -eq $wtJSON."$imp" -or $wtJSON."$imp" -eq $true)
            {
                $wtJSON | Add-Member -NotePropertyName $imp -NotePropertyValue $false -Force
            }
        }
    }

    # save settings
    $wtJSON | ConvertTo-Json -Depth 20 | Out-File "$wtAppData\settings.json" -Force -Encoding utf8
}


Write-Host -ForegroundColor Green "Please restart the console/terminal. Some changes will not take effect until after the restart."
