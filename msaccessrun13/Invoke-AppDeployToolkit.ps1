<#

.SYNOPSIS
PSAppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION
- Generic PSADT deployment script with flexible detection logic.
- Supports:
  - Installer types: EXE | MSI | MSIX | WINGET | SCRIPT
  - Detection modes: Uninstall | Exe | RegistryKey | RegistryValue | Custom
  - Detection handling: Internal | External | None
  - Post-install detection retry loop
  - SCCM / Intune external detection scenarios

.PARAMETER DeploymentType
The type of deployment to perform.

.PARAMETER DeployMode
Specifies whether the installation should be run in Interactive (shows dialogs), Silent (no dialogs), NonInteractive (dialogs without prompts) mode, or Auto.

.PARAMETER SuppressRebootPassThru
Suppresses the 3010 return code from being passed back to the parent process.

.PARAMETER TerminalServerMode
Changes to user install mode for RDS/Citrix servers.

.PARAMETER DisableLogging
Disables logging to file for the script.

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateSet('Install', 'Uninstall', 'Repair')]
    [string]$DeploymentType,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Auto', 'Interactive', 'NonInteractive', 'Silent')]
    [string]$DeployMode,

    [Parameter(Mandatory = $false)]
    [switch]$SuppressRebootPassThru,

    [Parameter(Mandatory = $false)]
    [switch]$TerminalServerMode,

    [Parameter(Mandatory = $false)]
    [switch]$DisableLogging
)

##================================================
## MARK: GLOBAL APP CONFIG
##================================================
$Global:Pkg = @{
    #================================================
    # Basic App Info
    #================================================

    # Metadata (shown in logs/UI)
    Vendor       = 'Microsoft'
    Name         = 'Microsoft Access Runtime 2013'
    Version      = '15.0.4569.1506'
    Arch         = 'x64'
    Lang         = 'de-DE'
    Revision     = '01'
    InstallTitle = 'Microsoft Access Runtime 2013 (x64) de-DE'

    #================================================
    # Installer
    #================================================

    # Installer type: EXE | MSI | MSIX | WINGET | SCRIPT
    InstallerType = 'EXE'

    # Files (relative to .\Files)
    # For Access Runtime 2013 use extracted Office setup instead of wrapper EXE
    InstallerFile   = 'setup.exe'
    TransformFile   = $null
    PatchFiles      = @()
    MsiProductCode  = $null


    # Arguments
    InstallArgs   = '/config ".\accessrt.ww\Config.xml"'
    UninstallType   = 'EXE'
    UninstallArgs = $null
    UninstallProductCode = $null

    MsiProperties = @('ALLUSERS=1', 'REBOOT=ReallySuppress')

    #================================================
    # SupportFiles Hooks (optional)
    #================================================

    # Optional PowerShell hook scripts (relative to .\Files)
    PreInstallHookFile    = $null
    PostInstallHookFile   = $null
    PreUninstallHookFile  = $null
    PostUninstallHookFile = $null

    #================================================
    # Detection Strategy
    #================================================

    # Internal = script validates install/uninstall state itself
    # External = SCCM / Intune handles detection
    # None     = no detection verification in script
    DetectionHandling = 'Internal'   # Internal | External | None

    # Detection mode when DetectionHandling = Internal
    DetectionMode = 'Uninstall'      # Uninstall | Exe | RegistryKey | RegistryValue | Custom

    # Uninstall detection
    DetectDisplayNameRegex = 'Microsoft Access Runtime.*2013'
    DetectMinVersion       = '15.0'

    # EXE detection
    DetectExePaths = @(
        @{
            Path       = "$env:ProgramFiles\Microsoft Office\Office15\MSACCESS.EXE"
            MinVersion = '15.0.4569.1503'
        },
        @{
            Path       = "${env:ProgramFiles(x86)}\Microsoft Office\Office15\MSACCESS.EXE"
            MinVersion = '15.0'
        },
        @{
            Path       = "$env:ProgramFiles\Microsoft Office\root\Office15\MSACCESS.EXE"
            MinVersion = '15.0'
        },
        @{
            Path       = "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15\MSACCESS.EXE"
            MinVersion = '15.0'
        }
    )

    # Registry key detection
    DetectRegistryKeyPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Office\15.0\Access\InstallRoot',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\15.0\Access\InstallRoot'
    )

    # Registry value detection
    # Example:
    # @{
    #   Path      = 'HKLM:\Software\Vendor\App'
    #   ValueName = 'InstallPath'
    #   MatchRegex = '.+'
    # }
    DetectRegistryValuePaths = @(
    )

    # Custom detection
    CustomDetectScriptBlock = $null

    # Detection retry window
    PostInstallDetectionDelay   = 10
    PostInstallDetectionRetries = 12

    #================================================
    # Process Handling
    #================================================

    # Processes to close (optional)
    ProcessesToClose = @('msaccess')

    #================================================
    # Exit Codes
    #================================================

    SuccessExitCodes = @(0)
    RebootExitCodes  = @(1641, 3010)

    #================================================
    # Optional: MSIX
    #================================================

    MsixAddArgs     = ''
    MsixRemoveArgs  = ''
    MsixPackageName = $null

    #================================================
    # Optional: Winget
    #================================================

    WingetId     = $null
    WingetScope  = 'machine'
    WingetSource = 'winget'
}

##================================================
## MARK: Variables
##================================================

$adtSession = @{
    AppVendor             = $Global:Pkg.Vendor
    AppName               = $Global:Pkg.Name
    AppVersion            = $Global:Pkg.Version
    AppArch               = $Global:Pkg.Arch
    AppLang               = $Global:Pkg.Lang
    AppRevision           = $Global:Pkg.Revision
    AppSuccessExitCodes   = $Global:Pkg.SuccessExitCodes
    AppRebootExitCodes    = $Global:Pkg.RebootExitCodes
    AppProcessesToClose   = $Global:Pkg.ProcessesToClose

    AppScriptVersion      = '1.1.0'
    AppScriptDate         = '2026-03-22'
    AppScriptAuthor       = 'Petar Zujovic'
    RequireAdmin          = $true

    InstallTitle          = $Global:Pkg.InstallTitle

    DeployAppScriptFriendlyName = $MyInvocation.MyCommand.Name
    DeployAppScriptParameters   = $PSBoundParameters
    DeployAppScriptVersion      = '4.1.8'
}

##================================================
## MARK: Helper Functions (Generic)
##================================================

function Get-UninstallEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DisplayNameRegex
    )

    $roots = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }

        Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue |
            ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue } |
            Where-Object { $_.DisplayName -and ($_.DisplayName -match $DisplayNameRegex) }
    }
}

function Get-UninstallEntry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DisplayNameRegex
    )

    $entries = @(Get-UninstallEntries -DisplayNameRegex $DisplayNameRegex)
    if (-not $entries -or $entries.Count -eq 0) { return $null }

    $withVersion = $entries | Where-Object { $_.DisplayVersion }
    if ($withVersion) {
        $sorted = $withVersion | Sort-Object -Property @{
            Expression = {
                try { [version]$_.DisplayVersion } catch { [version]'0.0.0.0' }
            }
        } -Descending
        return $sorted | Select-Object -First 1
    }

    return $entries | Select-Object -First 1
}

function Get-UninstallEntriesForPackageVersion {
    [CmdletBinding()]
    param()

    if (-not $Global:Pkg.DetectDisplayNameRegex) { return @() }

    $entries = Get-UninstallEntries -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex
    if (-not $entries) { return @() }

    $target = $Global:Pkg.Version
    $exact = $entries | Where-Object { $_.DisplayVersion -eq $target }
    if ($exact) { return $exact }

    return $entries
}

function Get-StandardMsiArgs {
    [CmdletBinding()]
    param(
        [ValidateSet('Install', 'Uninstall', 'Repair', 'Patch')]
        [string]$Action = 'Install'
    )

    $arguments = @('/qn', '/norestart')

    if ($Global:Pkg.MsiProperties -and $Global:Pkg.MsiProperties.Count -gt 0) {
        $arguments += ($Global:Pkg.MsiProperties -join ' ')
    }

    ($arguments -join ' ').Trim()
}

function Resolve-ConfiguredPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return $ExecutionContext.InvokeCommand.ExpandString($Path)
}

function Test-AppInstalledByUninstall {
    [CmdletBinding()]
    param()

    if (-not $Global:Pkg.DetectDisplayNameRegex) {
        Write-ADTLogEntry -Message "Detection(Uninstall): DetectDisplayNameRegex is empty." -Severity 2
        return $false
    }

    $entries = @(Get-UninstallEntries -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex)

    if (-not $entries -or $entries.Count -eq 0) {
        Write-ADTLogEntry -Message "Detection(Uninstall): No uninstall entry matched regex '$($Global:Pkg.DetectDisplayNameRegex)'." -Severity 1
        return $false
    }

    if ($Global:Pkg.DetectMinVersion) {
        $minVersion = [version]$Global:Pkg.DetectMinVersion

        foreach ($entry in $entries) {
            try {
                $entryVersion = [version]$entry.DisplayVersion
                if ($entryVersion -ge $minVersion) {
                    Write-ADTLogEntry -Message "Detection(Uninstall): Matched '$($entry.DisplayName)' version '$($entry.DisplayVersion)'." -Severity 1
                    return $true
                }
            }
            catch {
                Write-ADTLogEntry -Message "Detection(Uninstall): Failed to parse DisplayVersion '$($entry.DisplayVersion)' for '$($entry.DisplayName)'." -Severity 2
            }
        }

        Write-ADTLogEntry -Message "Detection(Uninstall): Matching entry found, but none met minimum version '$($Global:Pkg.DetectMinVersion)'." -Severity 1
        return $false
    }

    Write-ADTLogEntry -Message "Detection(Uninstall): Matched '$($entries[0].DisplayName)'." -Severity 1
    return $true
}

function Test-AppInstalledByExe {
    [CmdletBinding()]
    param()

    if (-not $Global:Pkg.DetectExePaths -or $Global:Pkg.DetectExePaths.Count -eq 0) {
        Write-ADTLogEntry -Message "Detection(Exe): DetectExePaths is empty." -Severity 2
        return $false
    }

    foreach ($item in $Global:Pkg.DetectExePaths) {
        $path = Resolve-ConfiguredPath -Path $item.Path

        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            continue
        }

        if ($item.MinVersion) {
            try {
                $file = Get-Item -LiteralPath $path -ErrorAction Stop
                $fileVersion = [version]$file.VersionInfo.ProductVersion
                $minVersion  = [version]$item.MinVersion

                if ($fileVersion -ge $minVersion) {
                    Write-ADTLogEntry -Message "Detection(Exe): Found '$path' version '$fileVersion'." -Severity 1
                    return $true
                }
            }
            catch {
                Write-ADTLogEntry -Message "Detection(Exe): Failed to evaluate version for '$path'. $($_.Exception.Message)" -Severity 2
            }
        }
        else {
            Write-ADTLogEntry -Message "Detection(Exe): Found '$path'." -Severity 1
            return $true
        }
    }

    Write-ADTLogEntry -Message "Detection(Exe): None of the DetectExePaths matched." -Severity 1
    return $false
}

function Test-AppInstalledByRegistryKey {
    [CmdletBinding()]
    param()

    if (-not $Global:Pkg.DetectRegistryKeyPaths -or $Global:Pkg.DetectRegistryKeyPaths.Count -eq 0) {
        Write-ADTLogEntry -Message "Detection(RegistryKey): DetectRegistryKeyPaths is empty." -Severity 2
        return $false
    }

    foreach ($path in $Global:Pkg.DetectRegistryKeyPaths) {
        $expandedPath = Resolve-ConfiguredPath -Path $path
        if (Test-Path -LiteralPath $expandedPath) {
            Write-ADTLogEntry -Message "Detection(RegistryKey): Registry key present: $expandedPath" -Severity 1
            return $true
        }
    }

    Write-ADTLogEntry -Message "Detection(RegistryKey): None of the DetectRegistryKeyPaths exist." -Severity 1
    return $false
}

function Test-AppInstalledByRegistryValue {
    [CmdletBinding()]
    param()

    if (-not $Global:Pkg.DetectRegistryValuePaths -or $Global:Pkg.DetectRegistryValuePaths.Count -eq 0) {
        Write-ADTLogEntry -Message "Detection(RegistryValue): DetectRegistryValuePaths is empty." -Severity 2
        return $false
    }

    foreach ($item in $Global:Pkg.DetectRegistryValuePaths) {
        try {
            $path = Resolve-ConfiguredPath -Path $item.Path
            $name = $item.ValueName
            $regex = $item.MatchRegex

            if (-not $path -or -not $name) { continue }
            if (-not (Test-Path -LiteralPath $path)) { continue }

            if ($name -eq '(default)') {
                $val = (Get-Item -LiteralPath $path -ErrorAction Stop).GetValue('')
            }
            else {
                $obj = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
                $val = $obj.$name
            }

            if ($null -eq $val) { continue }

            if ($regex) {
                if ("$val" -match $regex) {
                    Write-ADTLogEntry -Message "Detection(RegistryValue): '$path\$name' matched regex '$regex'. Value='$val'" -Severity 1
                    return $true
                }
            }
            else {
                if ("$val".Length -gt 0) {
                    Write-ADTLogEntry -Message "Detection(RegistryValue): '$path\$name' exists. Value='$val'" -Severity 1
                    return $true
                }
            }
        }
        catch {
            Write-ADTLogEntry -Message "Detection(RegistryValue): Failed reading registry value. $($_.Exception.Message)" -Severity 2
        }
    }

    Write-ADTLogEntry -Message "Detection(RegistryValue): None of the DetectRegistryValuePaths were satisfied." -Severity 1
    return $false
}

function Test-AppInstalled {
    [CmdletBinding()]
    param()

    switch ($Global:Pkg.DetectionMode) {
        'Uninstall'     { return Test-AppInstalledByUninstall }
        'Exe'           { return Test-AppInstalledByExe }
        'RegistryKey'   { return Test-AppInstalledByRegistryKey }
        'RegistryValue' { return Test-AppInstalledByRegistryValue }
        'Custom' {
            if ($Global:Pkg.CustomDetectScriptBlock) {
                return & $Global:Pkg.CustomDetectScriptBlock
            }
            throw "DetectionMode is 'Custom' but CustomDetectScriptBlock is not defined."
        }
        default {
            throw "Unknown DetectionMode '$($Global:Pkg.DetectionMode)'."
        }
    }
}

function Wait-ForAppDetection {
    [CmdletBinding()]
    param()

    $retries = if ($Global:Pkg.PostInstallDetectionRetries) { [int]$Global:Pkg.PostInstallDetectionRetries } else { 1 }
    $delay   = if ($Global:Pkg.PostInstallDetectionDelay)   { [int]$Global:Pkg.PostInstallDetectionDelay }   else { 0 }

    for ($i = 1; $i -le $retries; $i++) {
        if (Test-AppInstalled) {
            Write-ADTLogEntry -Message "Detection: Application detected on attempt $i of $retries." -Severity 1
            return $true
        }

        if ($i -lt $retries -and $delay -gt 0) {
            Write-ADTLogEntry -Message "Detection: Attempt $i of $retries failed. Waiting $delay second(s) before retry." -Severity 1
            Start-Sleep -Seconds $delay
        }
    }

    return $false
}

function Invoke-PostInstallDetectionPolicy {
    [CmdletBinding()]
    param()

    switch ($Global:Pkg.DetectionHandling) {
        'Internal' {
            if (-not (Wait-ForAppDetection)) {
                throw "Install completed but detection still fails for '$($Global:Pkg.Name)'."
            }
        }
        'External' {
            Write-ADTLogEntry -Message "Post-install detection is disabled in script. Detection is handled externally by SCCM / Intune." -Severity 1
        }
        'None' {
            Write-ADTLogEntry -Message "Post-install detection is disabled entirely." -Severity 1
        }
        default {
            throw "Unknown DetectionHandling value '$($Global:Pkg.DetectionHandling)'."
        }
    }
}

function Split-CommandLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CommandLine
    )

    $cmd = $CommandLine.Trim()
    $exe = $null
    $arguments = ''

    if ($cmd.StartsWith('"')) {
        $end = $cmd.IndexOf('"', 1)
        if ($end -gt 1) {
            $exe  = $cmd.Substring(1, $end - 1)
            $arguments = $cmd.Substring($end + 1).Trim()
        }
    }
    else {
        $split = $cmd.Split(' ', 2)
        $exe = $split[0]
        if ($split.Count -gt 1) { $arguments = $split[1] }
    }

    [pscustomobject]@{
        Exe  = $exe
        Args = $arguments
    }
}

##================================================
## MARK: Helper - Install
##================================================

function Invoke-InstallPayload {
    [CmdletBinding()]
    param()

    $type = $Global:Pkg.InstallerType.ToUpperInvariant()

    switch ($type) {
        'EXE' {
            $installer = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
                throw "Installer not found: $installer"
            }

            $result = Start-ADTProcess -FilePath $installer -ArgumentList $Global:Pkg.InstallArgs -NoWait:$false -PassThru
            if ($result.ExitCode -notin ($Global:Pkg.SuccessExitCodes + $Global:Pkg.RebootExitCodes)) {
                throw "Installer exited with code [$($result.ExitCode)]."
            }

            if ($result.ExitCode -in $Global:Pkg.RebootExitCodes) {
                Write-ADTLogEntry -Message "Installer completed with reboot exit code [$($result.ExitCode)]." -Severity 2
            }
            else {
                Write-ADTLogEntry -Message "Installer completed successfully with exit code [$($result.ExitCode)]." -Severity 1
            }
        }

        'MSI' {
            $msi = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $msi -PathType Leaf)) {
                throw "MSI not found: $msi"
            }

            $splat = @{
                Action       = 'Install'
                FilePath     = $msi
                ArgumentList = Get-StandardMsiArgs -Action Install
            }

            if ($Global:Pkg.TransformFile) {
                $mst = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.TransformFile)
                if (-not (Test-Path -LiteralPath $mst -PathType Leaf)) {
                    throw "MST not found: $mst"
                }
                $splat.Transforms = $mst
            }

            Start-ADTMsiProcess @splat

            if ($Global:Pkg.PatchFiles -and $Global:Pkg.PatchFiles.Count -gt 0) {
                foreach ($p in $Global:Pkg.PatchFiles) {
                    $msp = Join-Path $PSScriptRoot ("Files\" + $p)
                    if (-not (Test-Path -LiteralPath $msp -PathType Leaf)) {
                        throw "MSP not found: $msp"
                    }
                    Start-ADTMsiProcess -Action Patch -FilePath $msp -ArgumentList (Get-StandardMsiArgs -Action Patch)
                }
            }
        }

        'MSIX' {
            $pkgPath = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $pkgPath -PathType Leaf)) {
                throw "MSIX not found: $pkgPath"
            }

            Add-AppxPackage -Path $pkgPath -ErrorAction Stop
            Write-ADTLogEntry -Message "MSIX installation completed." -Severity 1
        }

        'WINGET' {
            if (-not $Global:Pkg.WingetId) {
                throw "WingetId is null. Set Global:Pkg.WingetId."
            }

            $winget = (Get-Command winget.exe -ErrorAction SilentlyContinue).Source
            if (-not $winget) {
                throw "winget.exe not found on this device."
            }

            $wingetArgs = @(
                'install', '--id', $Global:Pkg.WingetId,
                '--scope', $Global:Pkg.WingetScope,
                '--source', $Global:Pkg.WingetSource,
                '--silent',
                '--accept-package-agreements',
                '--accept-source-agreements'
            ) -join ' '

            $result = Start-ADTProcess -FilePath $winget -ArgumentList $wingetArgs -NoWait:$false -PassThru
            if ($result.ExitCode -ne 0) {
                throw "winget install exited with code [$($result.ExitCode)]."
            }

            Write-ADTLogEntry -Message "winget installation completed successfully." -Severity 1
        }

        'SCRIPT' {
            $payload = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $payload -PathType Leaf)) {
                throw "Script payload not found: $payload"
            }

            if ($payload.ToLowerInvariant().EndsWith('.ps1')) {
                $psArgs = "-ExecutionPolicy Bypass -NoProfile -File `"$payload`" $($Global:Pkg.InstallArgs)"
                $result = Start-ADTProcess -FilePath "powershell.exe" -ArgumentList $psArgs -NoWait:$false -PassThru
            }
            else {
                $result = Start-ADTProcess -FilePath $payload -ArgumentList $Global:Pkg.InstallArgs -NoWait:$false -PassThru
            }

            if ($result.ExitCode -ne 0) {
                throw "Script installer exited with code [$($result.ExitCode)]."
            }

            Write-ADTLogEntry -Message "Script installation completed successfully." -Severity 1
        }

        default {
            throw "Unsupported InstallerType '$type'. Use EXE|MSI|MSIX|WINGET|SCRIPT."
        }
    }
}

##================================================
## MARK: Helper Test MSI Product Code
##================================================

function Test-MsiProductCodeInstalled {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^\{[0-9A-Fa-f-]{36}\}$')]
        [string]$ProductCode
    )

    $paths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode"
    )

    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) { return $true }
    }

    $installerProductsRoot = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
    if (-not (Test-Path -LiteralPath $installerProductsRoot)) {
        return $false
    }

    function Convert-GuidToPackedProductCode {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidatePattern('^[0-9A-Fa-f]{32}$')]
            [string]$GuidNoHyphens
        )

        $p = $GuidNoHyphens.ToUpperInvariant()

        $a = $p.Substring(0, 8).ToCharArray()
        [array]::Reverse($a)
        $a = -join $a

        $b = $p.Substring(8, 4).ToCharArray()
        [array]::Reverse($b)
        $b = -join $b

        $c = $p.Substring(12, 4).ToCharArray()
        [array]::Reverse($c)
        $c = -join $c

        $d = $p.Substring(16, 16).ToCharArray()
        for ($i = 0; $i -lt $d.Length; $i += 2) {
            $tmp = $d[$i]
            $d[$i] = $d[$i + 1]
            $d[$i + 1] = $tmp
        }

        return "$a$b$c$(-join $d)"
    }

    $guidNoHyphens = $ProductCode.Trim('{}') -replace '-', ''
    if ($guidNoHyphens.Length -ne 32) { return $false }

    $packed = Convert-GuidToPackedProductCode -GuidNoHyphens $guidNoHyphens
    $packedKey = Join-Path $installerProductsRoot $packed

    return (Test-Path -LiteralPath $packedKey)
}

##================================================
## MARK: Helper Uninstall
##================================================

function Invoke-UninstallPayload {
    [CmdletBinding()]
    param()

    $type = if ($Global:Pkg.ContainsKey('UninstallType') -and $Global:Pkg.UninstallType) {
        $Global:Pkg.UninstallType.ToUpperInvariant()
    }
    else {
        $Global:Pkg.InstallerType.ToUpperInvariant()
    }

    switch ($type) {
        'EXE' {
            # Explicit uninstall
            if ($Global:Pkg.ContainsKey('UninstallFile') -and $Global:Pkg.UninstallFile) {
                $payload = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.UninstallFile)

                if (-not (Test-Path -LiteralPath $payload -PathType Leaf)) {
                    throw "Uninstall file not found: $payload"
                }

                Write-ADTLogEntry -Message "Uninstall: using explicit uninstall file [$payload]." -Severity 1

                $result = Start-ADTProcess -FilePath $payload -ArgumentList $Global:Pkg.UninstallArgs -NoWait:$false -PassThru
                if ($result.ExitCode -notin ($Global:Pkg.SuccessExitCodes + $Global:Pkg.RebootExitCodes)) {
                    throw "Uninstaller exited with code [$($result.ExitCode)]."
                }

                return
            }
            # Registry uninstall
            $entry = $null
            if ($Global:Pkg.DetectDisplayNameRegex) {
                $entry = Get-UninstallEntry -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex
            }

            if (-not $entry) {
                Write-ADTLogEntry -Message "Uninstall: no uninstall entry found." -Severity 2
                return
            }

            $cmd = $entry.QuietUninstallString
            if (-not $cmd) { $cmd = $entry.UninstallString }
            if (-not $cmd) { throw "Uninstall: uninstall command missing in registry." }

            $parsed = Split-CommandLine -CommandLine $cmd
            if (-not $parsed.Exe) { throw "Uninstall: failed to parse uninstall command: $cmd" }

            $arguments = $parsed.Args
            if ($Global:Pkg.UninstallArgs) {
                if ($arguments) { $arguments = "$arguments $($Global:Pkg.UninstallArgs)" }
                else { $arguments = $Global:Pkg.UninstallArgs }
            }

            $result = Start-ADTProcess -FilePath $parsed.Exe -ArgumentList $arguments -NoWait:$false -PassThru
            if ($result.ExitCode -notin ($Global:Pkg.SuccessExitCodes + $Global:Pkg.RebootExitCodes)) {
                throw "Uninstaller exited with code [$($result.ExitCode)]."
            }
        }

'MSI' {
    Write-ADTLogEntry -Message "MSI uninstall flow starting (multi-product)." -Severity 1

    $entries = @(Get-UninstallEntries -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex)

    if (-not $entries) {
        Write-ADTLogEntry -Message "No matching MSI entries found for uninstall." -Severity 2
        return
    }

    foreach ($entry in $entries) {
        $cmd = $entry.QuietUninstallString
        if (-not $cmd) { $cmd = $entry.UninstallString }

        if ($cmd -match '\{[0-9A-Fa-f-]{36}\}') {
            $guid = $Matches[0]

            Write-ADTLogEntry -Message "Uninstalling MSI ProductCode [$guid] ($($entry.DisplayName))" -Severity 1

            $splat = @{
                Action       = 'Uninstall'
                ProductCode  = $guid
                ArgumentList = '/qn /norestart'
            }

            Start-ADTMsiProcess @splat
        }
        else {
            Write-ADTLogEntry -Message "Skipping entry without valid ProductCode: $($entry.DisplayName)" -Severity 2
        }
    }

    Write-ADTLogEntry -Message "MSI uninstall flow completed for all detected products." -Severity 1
}

        'MSIX' {
            if ($Global:Pkg.MsixPackageName) {
                Get-AppxPackage -Name $Global:Pkg.MsixPackageName -AllUsers | Remove-AppxPackage -ErrorAction Stop
                return
            }

            $regex = $Global:Pkg.DetectDisplayNameRegex
            $pkg = Get-AppxPackage -AllUsers | Where-Object { $_.Name -match $regex -or $_.PackageFamilyName -match $regex } | Select-Object -First 1
            if (-not $pkg) {
                Write-ADTLogEntry -Message "Uninstall: no MSIX package matched." -Severity 2
                return
            }

            Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
        }

        'WINGET' {
            if (-not $Global:Pkg.WingetId) { throw "WingetId is null. Set Global:Pkg.WingetId." }

            $winget = (Get-Command winget.exe -ErrorAction SilentlyContinue).Source
            if (-not $winget) { throw "winget.exe not found on this device." }

            $wingetArgs = @(
                'uninstall', '--id', $Global:Pkg.WingetId,
                '--silent',
                '--accept-source-agreements'
            ) -join ' '

            $result = Start-ADTProcess -FilePath $winget -ArgumentList $wingetArgs -NoWait:$false -PassThru
            if ($result.ExitCode -ne 0) {
                throw "winget uninstall exited with code [$($result.ExitCode)]."
            }
        }

        'SCRIPT' {
            $payload = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $payload -PathType Leaf)) { throw "Uninstall script not found: $payload" }

            if ($payload.ToLowerInvariant().EndsWith('.ps1')) {
                $psArgs = "-ExecutionPolicy Bypass -NoProfile -File `"$payload`" $($Global:Pkg.UninstallArgs)"
                $result = Start-ADTProcess -FilePath "powershell.exe" -ArgumentList $psArgs -NoWait:$false -PassThru
            }
            else {
                $result = Start-ADTProcess -FilePath $payload -ArgumentList $Global:Pkg.UninstallArgs -NoWait:$false -PassThru
            }

            if ($result.ExitCode -ne 0) {
                throw "Script uninstall exited with code [$($result.ExitCode)]."
            }
        }

        default {
            throw "Unsupported InstallerType '$type'. Use EXE|MSI|MSIX|WINGET|SCRIPT."
        }
    }
}

##================================================
## MARK: Helper Stop Running App
##================================================

function Stop-RunningApp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$ProcessNames,

        [int]$GraceSeconds = 10,

        [int]$TaskkillTimeoutSeconds = 20
    )

    $names = foreach ($n in $ProcessNames) {
        $n = $n.Trim()
        if (-not $n) { continue }
        if ($n.ToLowerInvariant().EndsWith('.exe')) { $n.Substring(0, $n.Length - 4) } else { $n }
    }

    foreach ($name in $names) {
        Write-ADTLogEntry -Message "Process-close for '$name' (graceful then forced). DeployMode=$($adtSession.DeployMode)." -Severity 1

        $getProcs = {
            try { Get-Process -Name $name -ErrorAction Stop } catch { @() }
        }

        $procs = & $getProcs
        if ($procs.Count -gt 0) {
            foreach ($p in $procs) {
                try { $null = $p.CloseMainWindow() } catch {}
            }
            Start-Sleep -Seconds $GraceSeconds
        }
        else {
            Write-ADTLogEntry -Message "No running process '$name' found for graceful close." -Severity 1
        }

        $procs = & $getProcs
        if ($procs.Count -gt 0) {
            $taskkill = (Get-Command taskkill.exe -ErrorAction SilentlyContinue).Source
            if ($taskkill) {
                $arguments = "/F /IM `"$name.exe`" /T"

                try {
                    $p = Start-Process -FilePath $taskkill -ArgumentList $arguments -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop

                    if ($p.ExitCode -ne 0) {
                        Write-ADTLogEntry -Message "taskkill exit code $($p.ExitCode) for '$name'. Continuing (best-effort)." -Severity 2
                    }
                    else {
                        Write-ADTLogEntry -Message "taskkill succeeded for '$name'." -Severity 1
                    }
                }
                catch {
                    Write-ADTLogEntry -Message "taskkill invocation failed for '$name' (best-effort). $($_.Exception.Message)" -Severity 2
                }

                Start-Sleep -Seconds ([Math]::Min(2, $TaskkillTimeoutSeconds))
            }
            else {
                Write-ADTLogEntry -Message "taskkill.exe not found. Falling back to Stop-Process (in-session only)." -Severity 2
                foreach ($p in $procs) {
                    try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch {}
                }
            }
        }

        $still = & $getProcs
        if ($still.Count -gt 0) {
            $details = $still | Select-Object Id, ProcessName, SessionId | ForEach-Object { "PID=$($_.Id) Session=$($_.SessionId)" }
            Write-ADTLogEntry -Message "Process '$name' still running after close attempts: $($details -join '; ')" -Severity 2
        }
        else {
            Write-ADTLogEntry -Message "Process '$name' is not running." -Severity 1
        }
    }
}

##================================================
## MARK: Helper - SupportFiles
##================================================
function Invoke-SupportFileScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RelativePath
    )

    $supportRoot = Join-Path $PSScriptRoot 'SupportFiles'
    $scriptPath  = Join-Path $supportRoot $RelativePath

    if (-not (Test-Path -LiteralPath $scriptPath -PathType Leaf)) {
        throw "SupportFiles script not found: $scriptPath"
    }

    Write-ADTLogEntry -Message "Invoking SupportFiles script: $scriptPath" -Severity 1
    & $scriptPath
}

##================================================
## MARK: ADT Deployment Functions
##================================================

function Install-ADTDeployment {
    [CmdletBinding()]
    param()

    ##================================================
    ## MARK: Early detection (BEFORE UI)
    ##================================================
    if ($Global:Pkg.DetectionHandling -eq 'Internal') {
        if (Test-AppInstalled) {
            Write-ADTLogEntry -Message "$($Global:Pkg.Name) $($Global:Pkg.Version) already installed. Exiting." -Severity 1
            return
        }
    }
    elseif ($Global:Pkg.DetectionHandling -eq 'External') {
        Write-ADTLogEntry -Message "Early detection bypassed because DetectionHandling is External." -Severity 1
    }
    elseif ($Global:Pkg.DetectionHandling -eq 'None') {
        Write-ADTLogEntry -Message "Early detection bypassed because DetectionHandling is None." -Severity 1
    }

    ##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    if ($Global:Pkg.ContainsKey('PreInstallScript') -and $Global:Pkg.PreInstallScript) {
        Invoke-SupportFileScript -RelativePath $Global:Pkg.PreInstallScript
    }

    $welcome = @{
        CheckDiskSpace = $true
        PersistPrompt  = $true
    }

    if ($adtSession.AppProcessesToClose.Count -gt 0) {
        $welcome.CloseProcesses = $adtSession.AppProcessesToClose
        $welcome.CloseProcessesCountdown = 60
    }

    Show-ADTInstallationWelcome @welcome

    if ($adtSession.DeployMode -in @('Silent', 'NonInteractive', 'Auto')) {
        Stop-RunningApp -ProcessNames $Global:Pkg.ProcessesToClose -GraceSeconds 5
    }

    Show-ADTInstallationProgress

    ##================================================
    ## MARK: Install
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Invoke-InstallPayload
    Invoke-PostInstallDetectionPolicy

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

Write-ADTLogEntry -Message "PostInstallScript value: [$($Global:Pkg.PostInstallScript)]" -Severity 1

    if ($Global:Pkg.ContainsKey('PostInstallScript') -and $Global:Pkg.PostInstallScript) {
        Invoke-SupportFileScript -RelativePath $Global:Pkg.PostInstallScript
    }

    if ($adtSession.DeployMode -eq 'Interactive') {
        Show-ADTInstallationPrompt -Message "$($Global:Pkg.Name) installed." -ButtonRightText 'OK'
    }
}

function Uninstall-ADTDeployment {
    [CmdletBinding()]
    param()

    ##================================================
    ## MARK: Pre-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    if ($Global:Pkg.ContainsKey('PreUninstallScript') -and $Global:Pkg.PreUninstallScript) {
        Invoke-SupportFileScript -RelativePath $Global:Pkg.PreUninstallScript
    }

    $welcome = @{
        CheckDiskSpace = $false
        PersistPrompt  = $true
    }

    if ($adtSession.AppProcessesToClose -and $adtSession.AppProcessesToClose.Count -gt 0) {
        $welcome.CloseProcesses = $adtSession.AppProcessesToClose
        $welcome.CloseProcessesCountdown = 60
    }

    Show-ADTInstallationWelcome @welcome

    if ($adtSession.DeployMode -in @('Silent', 'NonInteractive', 'Auto')) {
        Stop-RunningApp -ProcessNames $Global:Pkg.ProcessesToClose -GraceSeconds 5
    }

    Show-ADTInstallationProgress

    ##================================================
    ## MARK: Uninstall
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Invoke-UninstallPayload

    if ($Global:Pkg.DetectionHandling -eq 'Internal') {
        if (Test-AppInstalled) {
            Write-ADTLogEntry -Message "Post-uninstall verification: '$($Global:Pkg.Name)' still appears installed." -Severity 2
        }
        else {
            Write-ADTLogEntry -Message "Post-uninstall verification: '$($Global:Pkg.Name)' is not installed." -Severity 1
        }
    }
    else {
        Write-ADTLogEntry -Message "Post-uninstall detection is skipped because DetectionHandling is '$($Global:Pkg.DetectionHandling)'." -Severity 1
    }

    Write-ADTLogEntry -Message "Uninstall flow finished. Returning success unless uninstaller threw an error." -Severity 1

    ##================================================
    ## MARK: Post-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    if ($Global:Pkg.ContainsKey('PostUninstallScript') -and $Global:Pkg.PostUninstallScript) {
        Invoke-SupportFileScript -RelativePath $Global:Pkg.PostUninstallScript
    }

}

function Repair-ADTDeployment {
    [CmdletBinding()]
    param()

    ##================================================
    ## MARK: Pre-Repair
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"
    Show-ADTInstallationProgress

    ##================================================
    ## MARK: Repair
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Invoke-InstallPayload

    if ($Global:Pkg.DetectionHandling -eq 'Internal') {
        if (-not (Wait-ForAppDetection)) {
            throw "Repair completed but detection still fails for '$($Global:Pkg.Name)'."
        }
    }
    elseif ($Global:Pkg.DetectionHandling -eq 'External') {
        Write-ADTLogEntry -Message "Repair completed. Detection is handled externally by SCCM / Intune." -Severity 1
    }
    else {
        Write-ADTLogEntry -Message "Repair completed. Detection is disabled." -Severity 1
    }

    ##================================================
    ## MARK: Post-Repair
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"
}

##================================================
## MARK: Initialization
##================================================

$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 1

try {
    if (Test-Path -LiteralPath "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1" -PathType Leaf) {
        Get-ChildItem -LiteralPath "$PSScriptRoot\PSAppDeployToolkit" -Recurse -File | Unblock-File -ErrorAction Ignore
        Import-Module -FullyQualifiedName @{
            ModuleName    = "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"
            Guid          = '8c3c366b-8606-4576-9f2d-4051144f7ca2'
            ModuleVersion = '4.1.8'
        } -Force
    }
    else {
        Import-Module -FullyQualifiedName @{
            ModuleName    = 'PSAppDeployToolkit'
            Guid          = '8c3c366b-8606-4576-9f2d-4051144f7ca2'
            ModuleVersion = '4.1.8'
        } -Force
    }

    $iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation
    $adtSession = Remove-ADTHashtableNullOrEmptyValues -Hashtable $adtSession
    $adtSession = Open-ADTSession @adtSession @iadtParams -PassThru
}
catch {
    $Host.UI.WriteErrorLine((Out-String -InputObject $_ -Width ([System.Int32]::MaxValue)))
    exit 60008
}

##================================================
## MARK: Invocation
##================================================

try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Directory | ForEach-Object {
        if ($_.Name -match 'PSAppDeployToolkit\..+$') {
            Get-ChildItem -LiteralPath $_.FullName -Recurse -File | Unblock-File -ErrorAction Ignore
            Import-Module -Name $_.FullName -Force
        }
    }

    & "$($adtSession.DeploymentType)-ADTDeployment"
    Close-ADTSession
}
catch {
    $mainErrorMessage = "An unhandled error within [$($MyInvocation.MyCommand.Name)] has occurred.`n$(Resolve-ADTErrorRecord -ErrorRecord $_)"
    Write-ADTLogEntry -Message $mainErrorMessage -Severity 3
    Close-ADTSession -ExitCode 60001
}