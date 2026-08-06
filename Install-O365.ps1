function Install-O365 {
    [CmdletBinding(DefaultParameterSetName = 'XMLFile')]
    param(
        [Parameter(ParameterSetName = 'XMLFile')][String]$ConfigurationXMLFile,
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('SemiAnnualPreview', 'SemiAnnual', 'MonthlyEnterprise', 'CurrentPreview', 'Current')]$Channel = 'Current',
        [Parameter(ParameterSetName = 'NoXML')][Switch]$DisplayInstall,
        [Parameter(ParameterSetName = 'NoXML')][Switch]$IncludeProject,
        [Parameter(ParameterSetName = 'NoXML')][Switch]$IncludeVisio,
        [Parameter(ParameterSetName = 'NoXML')][Array]$LanguageIDs,
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Groove', 'Outlook', 'OneNote', 'Access', 'OneDrive', 'Publisher', 'Word', 'Excel', 'PowerPoint', 'Teams', 'Lync')][Array]$ExcludeApps,
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('64', '32')]$OfficeArch = '64',
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('O365ProPlusRetail', 'O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet(0, 1)]$SharedComputerLicensing = '0',
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$EnableUpdates = 'TRUE',
        [Parameter(ParameterSetName = 'NoXML')][String]$SourcePath,
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$PinItemsToTaskbar = 'TRUE',
        [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$ForceOpenAppShutdown = 'FALSE',
        [Parameter(ParameterSetName = 'NoXML')][Switch]$KeepMSI,
        [Parameter(ParameterSetName = 'NoXML')][Switch]$RemoveAllProducts,
        [Parameter(ParameterSetName = 'NoXML')][Switch]$SetFileFormat,
        [Parameter(ParameterSetName = 'NoXML')][Switch]$ChangeArch,
        [Switch]$Uninstall = 'False',
        [String]$OfficeInstallDownloadPath = $(Join-Path $env:TEMP 'Office365Install'),
        [Switch]$LeaveInstallFiles = $False
    )

    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'

    $CurrentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (!($CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
        Write-Warning 'Script is not running as Administrator'
        Write-Warning 'Please rerun this script as Administrator.'
        exit
    }

    function Get-XMLFile {

        [CmdletBinding(DefaultParameterSetName = 'NoXML')]
        param(
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('SemiAnnualPreview', 'SemiAnnual', 'MonthlyEnterprise', 'CurrentPreview', 'Current')]$Channel = 'Current',
            [Parameter(ParameterSetName = 'NoXML')][Switch]$DisplayInstall,
            [Parameter(ParameterSetName = 'NoXML')][Switch]$IncludeProject,
            [Parameter(ParameterSetName = 'NoXML')][Switch]$IncludeVisio,
            [Parameter(ParameterSetName = 'NoXML')][Array]$LanguageIDs,
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('Groove', 'Outlook', 'OneNote', 'Access', 'OneDrive', 'Publisher', 'Word', 'Excel', 'PowerPoint', 'Teams', 'Lync')][Array]$ExcludeApps = @('Groove', 'Lync'),
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('64', '32')]$OfficeArch = '64',
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('O365ProPlusRetail', 'O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet(0, 1)]$SharedComputerLicensing = '0',
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$EnableUpdates = 'TRUE',
            [Parameter(ParameterSetName = 'NoXML')][String]$SourcePath,
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$PinItemsToTaskbar = 'TRUE',
            [Parameter(ParameterSetName = 'NoXML')][ValidateSet('TRUE', 'FALSE')]$ForceOpenAppShutdown = 'FALSE',
            [Parameter(ParameterSetName = 'NoXML')][Switch]$KeepMSI,
            [Parameter(ParameterSetName = 'NoXML')][Switch]$RemoveAllProducts,
            [Parameter(ParameterSetName = 'NoXML')][Switch]$SetFileFormat,
            [Parameter(ParameterSetName = 'NoXML')][Switch]$ChangeArch
        )

        if ($ExcludeApps) {
            $ExcludeApps | ForEach-Object {
                $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
            }
        }

        if ($LanguageIDs) {
            $LanguageIDs | ForEach-Object {
                $LanguageString += "<Language ID =`"$_`" />"
            }
        } else {
            $LanguageString = "<Language ID=`"MatchOS`" />"
        }

        if ($OfficeArch) {
            $OfficeArchString = "`"$OfficeArch`""
        }

        if ($ChangeArch) {
            $MigrateArch = "MigrateArch=`"TRUE`""
        } else {
            $MigrateArch = $Null
        }

        if ($KeepMSI) {
            $RemoveMSIString = $Null
        } else {
            $RemoveMSIString = '<RemoveMSI />'
        }

        if ($RemoveAllProducts) {
            $RemoveAllString = "<Remove All=`"TRUE`" />"
        } else {
            $RemoveAllString = $Null
        }

        if ($SetFileFormat) {
            $AppSettingsString = '<AppSettings>
      <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
      <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
      <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
    </AppSettings>'
        } else {
            $AppSettingsString = $Null
        }

        if ($Channel) {
            $ChannelString = "Channel=`"$Channel`""
        } else {
            $ChannelString = $Null
        }

        if ($SourcePath) {
            $SourcePathString = "SourcePath=`"$SourcePath`""
        } else {
            $SourcePathString = $Null
        }

        if ($DisplayInstall) {
            $SilentInstallString = 'Full'
        } else {
            $SilentInstallString = 'None'
        }

        if ($IncludeProject) {
            $ProjectString = "<Product ID=`"ProjectProRetail`"`>$ExcludeAppsString $LanguageString</Product>"
        } else {
            $ProjectString = $Null
        }

        if ($IncludeVisio) {
            $VisioString = "<Product ID=`"VisioProRetail`"`>$ExcludeAppsString $LanguageString</Product>"
        } else {
            $VisioString = $Null
        }

        if ($Uninstall) {
            $OfficeXML = [XML]@"
  <Configuration>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Remove All="TRUE">
        <Product ID="$OfficeEdition" />
    </Remove>
    <RemoveMSI />
  </Configuration>
"@
        } else {
            $OfficeXML = [XML]@"
  <Configuration>
    <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString $MigrateArch >
      <Product ID="$OfficeEdition">
        $LanguageString
        $ExcludeAppsString
      </Product>
      $ProjectString
      $VisioString
    </Add>
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="FORCEAPPSHUTDOWN" Value="$ForceOpenAppShutdown" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Updates Enabled="$EnableUpdates" />
    $AppSettingsString
    $RemoveMSIString
    $RemoveAllString
  </Configuration>
"@
        }

        $OfficeXML
 
    }

    function Get-ODTURL {

        $OfficeDeploymentRegex = '"url":"(https:\/\/download\.microsoft\.com\/[^"]*officedeploymenttool[^"]*)"'

        [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'

        $MSWebPage | ForEach-Object {
            if ($_ -match $OfficeDeploymentRegex) {
                $matches[1]
            }
        }

    }

    if (-Not(Test-Path $OfficeInstallDownloadPath )) {
        New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
    }

    if (!($ConfigurationXMLFile)) {

        if ($ExcludeApps) {
            $OfficeXML = Get-XMLFile -AcceptEULA $AcceptEULA `
                -Channel $Channel `
                -DisplayInstall:$DisplayInstall `
                -IncludeProject:$IncludeProject `
                -IncludeVisio:$IncludeVisio `
                -LanguageIDs $LanguageIDs `
                -ExcludeApps $ExcludeApps `
                -OfficeArch $OfficeArch `
                -OfficeEdition $OfficeEdition `
                -SharedComputerLicensing $SharedComputerLicensing `
                -EnableUpdate $EnableUpdates `
                -PinItemsToTaskBar $PinItemsToTaskbar `
                -ForceOpenAppShutdown $ForceOpenAppShutdown `
                -KeepMSI:$KeepMSI `
                -RemoveAllProducts:$RemoveAllProducts `
                -SetFileFormat:$SetFileFormat `
                -ChangeArch:$ChangeArch `
 
        } else {
            $OfficeXML = Get-XMLFile -AcceptEULA $AcceptEULA `
                -Channel $Channel `
                -DisplayInstall:$DisplayInstall `
                -IncludeProject:$IncludeProject `
                -IncludeVisio:$IncludeVisio `
                -LanguageIDs $LanguageIDs `
                -OfficeArch $OfficeArch `
                -OfficeEdition $OfficeEdition `
                -SharedComputerLicensing $SharedComputerLicensing `
                -EnableUpdate $EnableUpdates `
                -PinItemsToTaskBar $PinItemsToTaskbar `
                -ForceOpenAppShutdown $ForceOpenAppShutdown `
                -KeepMSI:$KeepMSI `
                -RemoveAllProducts:$RemoveAllProducts `
                -SetFileFormat:$SetFileFormat `
                -ChangeArch:$ChangeArch `
 
        }

        $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")

        $ConfigurationXMLFile = "$OfficeInstallDownloadPath\OfficeInstall.xml"
    } else {

        if (!(Test-Path $ConfigurationXMLFile)) {
            Write-Warning 'The configuration XML file is not a valid file'
            Write-Warning 'Please check the path and try again'
            exit
        }

    }

    $ODTInstallLink = Get-ODTURL

    if ($Null -eq $ODTInstallLink) {
        Write-Error "Could not find ODT install link, exiting"
        exit
    }

    #Download the Office Deployment Tool
    Write-Verbose 'Downloading the Office Deployment Tool...'
    try {
        Invoke-WebRequest -Uri $ODTInstallLink -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
    } catch {
        Write-Warning 'There was an error downloading the Office Deployment Tool.'
        Write-Warning 'Please verify the below link is valid:'
        Write-Warning $ODTInstallLink
        exit
    }

    #Run the Office Deployment Tool setup
    try {
        Write-Verbose 'Running the Office Deployment Tool...'
        Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:`"$OfficeInstallDownloadPath`"" -Wait
    } catch {
        Write-Warning 'Error running the Office Deployment Tool. The error is below:'
        Write-Warning $_
    }

    #Run the O365 install
    try {
        Write-Verbose 'Downloading and Configuring Microsoft 365'
        $__ = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure `"$ConfigurationXMLFile`"" -Wait -PassThru -WindowStyle Hidden
    } catch {
        Write-Warning 'Error running the Office install. The error is below:'
        Write-Warning $_
    }

    #Check if Office 365 suite was installed correctly.
    $RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $OfficeInstalled = $False
    foreach ($Key in (Get-ChildItem $RegLocations) ) {
        if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
            $OfficeVersionInstalled = $Key.GetValue('DisplayName')
            $OfficeInstalled = $True
        }
    }

    if ($OfficeInstalled) {
        Write-Verbose "$($OfficeVersionInstalled) installed successfully!"
    } elseif ($Uninstall) {
        Write-Verbose 'Microsoft 365 uninstalled successfully!'
    } else {
        Write-Warning 'Microsoft 365 was not detected after the install ran'
    }

    if (!$LeaveInstallFiles) {
        Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
    }
}

if ($Arguments -ne "") {
    Invoke-Expression "Install-O365 $Arguments"
} else {
    Install-O365
}