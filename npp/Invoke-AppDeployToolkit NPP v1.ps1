<#

.SYNOPSIS
PSAppDeployToolkit - This script performs the installation or uninstallation of an application(s).

.DESCRIPTION
- The script is provided as a template to perform an install, uninstall, or repair of an application(s).
- The script either performs an "Install", "Uninstall", or "Repair" deployment type.
- The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.

The script imports the PSAppDeployToolkit module which contains the logic and functions required to install or uninstall an application.

.PARAMETER DeploymentType
The type of deployment to perform.

.PARAMETER DeployMode
Specifies whether the installation should be run in Interactive (shows dialogs), Silent (no dialogs), NonInteractive (dialogs without prompts) mode, or Auto (shows dialogs if a user is logged on, device is not in the OOBE, and there's no running apps to close).

Silent mode is automatically set if it is detected that the process is not user interactive, no users are logged on, the device is in Autopilot mode, or there's specified processes to close that are currently running.

.PARAMETER SuppressRebootPassThru
Suppresses the 3010 return code (requires restart) from being passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.

.PARAMETER TerminalServerMode
Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Desktop Session Hosts/Citrix servers.

.PARAMETER DisableLogging
Disables logging to file for the script.

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -DeployMode Silent

.EXAMPLE
powershell.exe -File Invoke-AppDeployToolkit.ps1 -DeploymentType Uninstall

.EXAMPLE
Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None. This script does not generate any output.

.NOTES
Toolkit Exit Code Ranges:
- 60000 - 68999: Reserved for built-in exit codes in Invoke-AppDeployToolkit.ps1, and Invoke-AppDeployToolkit.exe
- 69000 - 69999: Recommended for user customized exit codes in Invoke-AppDeployToolkit.ps1
- 70000 - 79999: Recommended for user customized exit codes in PSAppDeployToolkit.Extensions module.

.LINK
https://psappdeploytoolkit.com

Folder layout:
.\Invoke-AppDeployToolkit.ps1 (this file)
.\PSAppDeployToolkit\...
.\Files\  (installers go here)

Examples:
- EXE:  Files\7z2600-x64.exe
- MSI:  Files\App.msi (+ optional .mst)
- MSIX: Files\App.msix (or .appx/.appxbundle/.msixbundle)
- Script: Files\install.cmd / install.ps1

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
    # Metadata (shown in logs/UI)
    Vendor      = 'Notepad++ Team'
    Name        = 'Notepad++'
    Version     = '8.9.1'
    Arch        = 'x64'
    Lang        = 'DE'
    Revision    = '01'
    InstallTitle = 'Notepad++ 8.9.1 (x64)'

    # Installer type: EXE | MSI | MSIX | WINGET | SCRIPT
    InstallerType = 'MSI'

    # Files (relative to .\Files)
    InstallerFile = 'npp.8.9.1.Installer.x64.msi'                # EXE/MSI/MSIX/SCRIPT
    TransformFile = $null                                        # MSI only (.mst) optional
    PatchFiles    = @()                                          # MSI only (.msp) optional list
    MsiProductCode = '{7349B4F3-02E1-4234-A67A-FA85B33B67AF}'    # Optional MSI ProductCode GUID

    # Arguments
    InstallArgs   = $null               # EXE/SCRIPT/WINGET = '/S' | else = $null
    UninstallArgs = $null               # EXE uninstaller extra args (if needed) = '/S' | else = $null
    MsiProperties = @('ALLUSERS=1','REBOOT=ReallySuppress') # MSI only: additional properties to add to msiexec command line (e.g. 'ALLUSERS=1 REBOOT=ReallySuppress')


    # Detection: registry display name match + minimum version
    DetectDisplayNameRegex = '^Notepad\+\+\b'
    DetectMinVersion       = '8.9.1'

    # Fallback detection: file version check (recommended)
    DetectExePaths = @(
        "$env:ProgramFiles\Notepad++\notepad++.exe",
        "$env:ProgramFiles(x86)\Notepad++\notepad++.exe"
    )

    # Processes to close (optional)
    ProcessesToClose = @('notepad++')  # Example: @('excel', @{ Name = 'winword'; Description = 'Microsoft Word' })

    # Exit codes
    SuccessExitCodes = @(0)
    RebootExitCodes  = @(1641, 3010)

    # MSIX options
    MsixAddArgs      = ''               # e.g. '-ForceApplicationShutdown'
    MsixRemoveArgs   = ''               # e.g. '-AllUsers'
    MsixPackageName  = $null            # optional explicit package family/name if you prefer remove by name

    # Winget options (requires winget present)
    WingetId         = $null            # e.g. '7zip.7zip'
    WingetScope      = 'machine'        # machine|user
    WingetSource     = 'winget'
}


##================================================
## MARK: Variables
##================================================

# Zero-Config MSI support is provided when "AppName" is null or empty.
# By setting the "AppName" property, Zero-Config MSI will be disabled.
$adtSession = @{
    # App variables.
    AppVendor  = $Global:Pkg.Vendor
    AppName    = $Global:Pkg.Name
    AppVersion = $Global:Pkg.Version
    AppArch    = $Global:Pkg.Arch
    AppLang    = $Global:Pkg.Lang
    AppRevision = $Global:Pkg.Revision
    AppSuccessExitCodes = $Global:Pkg.SuccessExitCodes
    AppRebootExitCodes  = $Global:Pkg.RebootExitCodes
    AppProcessesToClose = $Global:Pkg.ProcessesToClose

    AppScriptVersion = '1.0.0'
    AppScriptDate    = '2026-02-15'
    AppScriptAuthor  = 'Petar Zujovic'
    RequireAdmin     = $true

    InstallTitle = $Global:Pkg.InstallTitle

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

    # Prefer entries with DisplayVersion that can be parsed, then take highest version
    $withVersion = $entries | Where-Object { $_.DisplayVersion }
    if ($withVersion) {
        $sorted = $withVersion | Sort-Object -Property @{
            Expression = {
                try { [version]$_.DisplayVersion } catch { [version]'0.0.0.0' }
            }
        } -Descending
        return $sorted | Select-Object -First 1
    }

    # Fallback: no DisplayVersion anywhere
    return $entries | Select-Object -First 1
}

function Get-UninstallEntriesForPackageVersion {
    [CmdletBinding()]
    param()

    $entries = Get-UninstallEntries -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex
    if (-not $entries) { return @() }

    $target = $Global:Pkg.Version

    # Prefer exact DisplayVersion match when available
    $exact = $entries | Where-Object { $_.DisplayVersion -eq $target }
    if ($exact) { return $exact }

    # Fallback: if DisplayVersion isn't present/consistent, return all matches
    return $entries
}

function Get-StandardMsiArgs {
    [CmdletBinding()]
    param(
        [ValidateSet('Install','Uninstall','Repair','Patch')]
        [string]$Action = 'Install'
    )

    # UI + reboot behavior
    # /qn = silent, /norestart = do not reboot automatically
    $args = @('/qn', '/norestart')

    # Add extra MSI properties (KEY=VALUE) if defined
    if ($Global:Pkg.MsiProperties -and $Global:Pkg.MsiProperties.Count -gt 0) {
        $args += ($Global:Pkg.MsiProperties -join ' ')
    }

    ($args -join ' ').Trim()
}


function Test-AppInstalled {
    [CmdletBinding()]
    param()

    # 1) Registry-based detection
    $entry = Get-UninstallEntry -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex
    if ($entry) {
        Write-ADTLogEntry -Message "Detection: Uninstall entry found: $($entry.DisplayName) $($entry.DisplayVersion)" -Severity 1
        if ($entry.DisplayVersion) {
            try {
                if ([version]$entry.DisplayVersion -ge [version]$Global:Pkg.DetectMinVersion) { return $true }
            } catch {
                return $true
            }
        }
    } else {
        Write-ADTLogEntry -Message "Detection: No uninstall entry matched regex '$($Global:Pkg.DetectDisplayNameRegex)'." -Severity 1
    }

    # 2) File-version fallback (prevents reinstall loops if DisplayVersion is blank)
    foreach ($p in $Global:Pkg.DetectExePaths) {
        if (Test-Path -LiteralPath $p -PathType Leaf) {
            $fv = (Get-Item -LiteralPath $p).VersionInfo.FileVersion
            Write-ADTLogEntry -Message "Detection: Found file '$p' version '$fv'." -Severity 1
            if (-not $fv) { return $true }
            try {
                if ([version]$fv -ge [version]$Global:Pkg.DetectMinVersion) { return $true }
            } catch {
                return $true
            }
        }
    }

    return $false
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
    } else {
        $split = $cmd.Split(' ', 2)
        $exe = $split[0]
        if ($split.Count -gt 1) { $arguments = $split[1] }
    }

    [pscustomobject]@{ Exe = $exe; Args = $arguments }
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
            if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) { throw "Installer not found: $installer" }
            Start-ADTProcess -FilePath $installer -ArgumentList $Global:Pkg.InstallArgs -NoWait:$false
        }

        'MSI' {
            $msi = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $msi -PathType Leaf)) { throw "MSI not found: $msi" }

            $splat = @{
                Action   = 'Install'
                FilePath = $msi
            }

            if ($Global:Pkg.TransformFile) {
                $mst = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.TransformFile)
                if (-not (Test-Path -LiteralPath $mst -PathType Leaf)) { throw "MST not found: $mst" }
                $splat.Transforms = $mst
            }

            # add /qn and /norestart (and your properties)
            $splat.ArgumentList = Get-StandardMsiArgs -Action Install

            Start-ADTMsiProcess @splat

            # PATCH SUPPORT
            if ($Global:Pkg.PatchFiles -and $Global:Pkg.PatchFiles.Count -gt 0) {
                foreach ($p in $Global:Pkg.PatchFiles) {
                    $msp = Join-Path $PSScriptRoot ("Files\" + $p)
                    if (-not (Test-Path -LiteralPath $msp -PathType Leaf)) { throw "MSP not found: $msp" }
                    Start-ADTMsiProcess -Action Patch -FilePath $msp -ArgumentList (Get-StandardMsiArgs -Action Patch)
                }
            }
        }


        'MSIX' {
            $pkgPath = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $pkgPath -PathType Leaf)) { throw "MSIX not found: $pkgPath" }
            # Requires running as admin. Uses built-in cmdlet.
            $arguments = $Global:Pkg.MsixAddArgs
            if ($arguments) {
                # If you need complex args, edit here. Keeping it simple.
                Add-AppxPackage -Path $pkgPath -ErrorAction Stop
            } else {
                Add-AppxPackage -Path $pkgPath -ErrorAction Stop
            }
        }

        'WINGET' {
            if (-not $Global:Pkg.WingetId) { throw "WingetId is null. Set Global:Pkg.WingetId." }
            $id = $Global:Pkg.WingetId
            $scope = $Global:Pkg.WingetScope
            $source = $Global:Pkg.WingetSource

            $winget = (Get-Command winget.exe -ErrorAction SilentlyContinue).Source
            if (-not $winget) { throw "winget.exe not found on this device." }

            $wingetArgs = @(
                'install', '--id', $id,
                '--scope', $scope,
                '--source', $source,
                '--silent',
                '--accept-package-agreements',
                '--accept-source-agreements'
            ) -join ' '

            Start-ADTProcess -FilePath $winget -ArgumentList $wingetArgs -NoWait:$false
        }

        'SCRIPT' {
            $payload = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $payload -PathType Leaf)) { throw "Script payload not found: $payload" }

            # If payload is .ps1, run with powershell; otherwise execute directly.
            if ($payload.ToLowerInvariant().EndsWith('.ps1')) {
                $psArgs = "-ExecutionPolicy Bypass -NoProfile -File `"$payload`" $($Global:Pkg.InstallArgs)"
                Start-ADTProcess -FilePath "powershell.exe" -ArgumentList $psArgs -NoWait:$false
            } else {
                Start-ADTProcess -FilePath $payload -ArgumentList $Global:Pkg.InstallArgs -NoWait:$false
            }
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

    # 1) Standard uninstall registry keys
    $paths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ProductCode"
    )
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) { return $true }
    }

    # 2) Windows Installer registration
    $installerProductsRoot = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
    if (-not (Test-Path -LiteralPath $installerProductsRoot)) {
        return $false
    }

    # MSI stores ProductCode GUIDs in a "packed" form under Installer\UserData\...\Products
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

    $type = $Global:Pkg.InstallerType.ToUpperInvariant()

    switch ($type) {
        'EXE' {
            $entry = Get-UninstallEntry -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex
            if (-not $entry) { Write-ADTLogEntry -Message "Uninstall: no uninstall entry found." -Severity 2; return }

            $cmd = $entry.QuietUninstallString
            if (-not $cmd) { $cmd = $entry.UninstallString }
            if (-not $cmd) { throw "Uninstall: uninstall command missing in registry." }

            $parsed = Split-CommandLine -CommandLine $cmd
            if (-not $parsed.Exe) { throw "Uninstall: failed to parse uninstall command: $cmd" }

            $arguments = $parsed.Args
            # Ensure silent uninstall (many EXE uninstallers accept /S)
            if ($Global:Pkg.UninstallArgs) {
                if ($arguments) { $arguments = "$arguments $($Global:Pkg.UninstallArgs)" } else { $arguments = $Global:Pkg.UninstallArgs }
            }
            Start-ADTProcess -FilePath $parsed.Exe -ArgumentList $arguments -NoWait:$false
        }

        'MSI' {

        Write-ADTLogEntry -Message "MSI uninstall flow starting." -Severity 1

        # -------------------------------------------------
        # Preferred method: use configured ProductCode
        # -------------------------------------------------
        if ($Global:Pkg.MsiProductCode -and (Test-MsiProductCodeInstalled -ProductCode $Global:Pkg.MsiProductCode)) {

            Write-ADTLogEntry -Message "Uninstall: removing MSI using configured ProductCode $($Global:Pkg.MsiProductCode)" -Severity 1

            $splat = @{
                Action       = 'Uninstall'
                ProductCode  = $Global:Pkg.MsiProductCode
                ArgumentList = Get-StandardMsiArgs -Action Uninstall
            }

            Write-ADTLogEntry -Message "Executing MSI uninstall with args: $($splat.ArgumentList)" -Severity 1

            Start-ADTMsiProcess @splat

            Write-ADTLogEntry -Message "MSI uninstall completed using configured ProductCode." -Severity 1

            return
        }

        # -------------------------------------------------
        # Fallback method: discover ProductCode dynamically
        # -------------------------------------------------
        Write-ADTLogEntry -Message "Configured ProductCode not found or not installed. Attempting registry discovery." -Severity 2

        $entry = Get-UninstallEntry -DisplayNameRegex $Global:Pkg.DetectDisplayNameRegex

        if (-not $entry) {

            Write-ADTLogEntry -Message "Uninstall: no MSI uninstall entry found matching '$($Global:Pkg.DetectDisplayNameRegex)'." -Severity 2

            return
        }

        $cmd = $entry.QuietUninstallString

        if (-not $cmd) {
            $cmd = $entry.UninstallString
        }

        if (-not $cmd) {

            Write-ADTLogEntry -Message "Uninstall: uninstall string missing in registry." -Severity 3

            throw "MSI uninstall failed: uninstall string missing."
        }

        if ($cmd -match '\{[0-9A-Fa-f-]{36}\}') {

            $guid = $Matches[0]

            Write-ADTLogEntry -Message "Discovered MSI ProductCode from registry: $guid" -Severity 1

            $splat = @{
                Action       = 'Uninstall'
                ProductCode  = $guid
                ArgumentList = Get-StandardMsiArgs -Action Uninstall
            }

            Write-ADTLogEntry -Message "Executing MSI uninstall with args: $($splat.ArgumentList)" -Severity 1

            Start-ADTMsiProcess @splat

            Write-ADTLogEntry -Message "MSI uninstall completed using discovered ProductCode." -Severity 1

        }
        else {

            Write-ADTLogEntry -Message "Uninstall string does not contain a valid MSI ProductCode: $cmd" -Severity 3

            throw "MSI uninstall failed: could not extract ProductCode."
        }

        Write-ADTLogEntry -Message "MSI uninstall flow finished successfully." -Severity 1
    }



        'MSIX' {
            # Remove by explicit package name if provided; else attempt by regex match from installed packages
            if ($Global:Pkg.MsixPackageName) {
                Get-AppxPackage -Name $Global:Pkg.MsixPackageName -AllUsers | Remove-AppxPackage -ErrorAction Stop
                return
            }

            $regex = $Global:Pkg.DetectDisplayNameRegex
            $pkg = Get-AppxPackage -AllUsers | Where-Object { $_.Name -match $regex -or $_.PackageFamilyName -match $regex } | Select-Object -First 1
            if (-not $pkg) { Write-ADTLogEntry -Message "Uninstall: no MSIX package matched." -Severity 2; return }
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

            Start-ADTProcess -FilePath $winget -ArgumentList $wingetArgs -NoWait:$false
        }

        'SCRIPT' {
            # If you have a dedicated uninstall script, set InstallerFile to it when DeploymentType=Uninstall
            $payload = Join-Path $PSScriptRoot ("Files\" + $Global:Pkg.InstallerFile)
            if (-not (Test-Path -LiteralPath $payload -PathType Leaf)) { throw "Uninstall script not found: $payload" }

            if ($payload.ToLowerInvariant().EndsWith('.ps1')) {
                $psArgs = "-ExecutionPolicy Bypass -NoProfile -File `"$payload`" $($Global:Pkg.UninstallArgs)"
                Start-ADTProcess -FilePath "powershell.exe" -ArgumentList $psArgs -NoWait:$false
            } else {
                Start-ADTProcess -FilePath $payload -ArgumentList $Global:Pkg.UninstallArgs -NoWait:$false
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

    # Normalize names: remove .exe if present
    $names = foreach ($n in $ProcessNames) {
        $n = $n.Trim()
        if (-not $n) { continue }
        if ($n.ToLowerInvariant().EndsWith('.exe')) { $n.Substring(0, $n.Length - 4) } else { $n }
    }

    foreach ($name in $names) {
        Write-ADTLogEntry -Message "Process-close for '$name' (graceful then forced). DeployMode=$($adtSession.DeployMode)." -Severity 1

        # Helper: get processes (best-effort)
        $getProcs = {
            try { Get-Process -Name $name -ErrorAction Stop } catch { @() }
        }

        # 1) Graceful close (in-session only)
        $procs = & $getProcs
        if ($procs.Count -gt 0) {
            foreach ($p in $procs) {
                try { $null = $p.CloseMainWindow() } catch {}
            }
            Start-Sleep -Seconds $GraceSeconds
        } else {
            Write-ADTLogEntry -Message "No running process '$name' found for graceful close." -Severity 1
        }

        # 2) Forced close (cross-session) via taskkill best-effort
        $procs = & $getProcs
        if ($procs.Count -gt 0) {
            $taskkill = (Get-Command taskkill.exe -ErrorAction SilentlyContinue).Source
            if ($taskkill) {
                $args = "/F /IM `"$name.exe`" /T"

                try {
                    # Use Start-Process so we can treat non-zero as non-fatal
                    $p = Start-Process -FilePath $taskkill -ArgumentList $args -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop

                    # taskkill uses various non-zero exit codes for "not found" / "no instance",
                    # but at this point we already saw it running, so any non-zero is "warn", not "fail".
                    if ($p.ExitCode -ne 0) {
                        Write-ADTLogEntry -Message "taskkill exit code $($p.ExitCode) for '$name'. Continuing (best-effort)." -Severity 2
                    } else {
                        Write-ADTLogEntry -Message "taskkill succeeded for '$name'." -Severity 1
                    }
                }
                catch {
                    Write-ADTLogEntry -Message "taskkill invocation failed for '$name' (best-effort). $($_.Exception.Message)" -Severity 2
                }

                # Give Windows a moment to tear down child processes
                Start-Sleep -Seconds ([Math]::Min(2, $TaskkillTimeoutSeconds))
            }
            else {
                Write-ADTLogEntry -Message "taskkill.exe not found. Falling back to Stop-Process (in-session only)." -Severity 2
                foreach ($p in $procs) {
                    try { Stop-Process -Id $p.Id -Force -ErrorAction Stop } catch {}
                }
            }
        }

        # 3) Verify
        $still = & $getProcs
        if ($still.Count -gt 0) {
            $details = $still | Select-Object Id, ProcessName, SessionId | ForEach-Object { "PID=$($_.Id) Session=$($_.SessionId)" }
            Write-ADTLogEntry -Message "Process '$name' still running after close attempts: $($details -join '; ')" -Severity 2
        } else {
            Write-ADTLogEntry -Message "Process '$name' is not running." -Severity 1
        }
    }
}

##================================================
## MARK: ADT Deployment Functions
##================================================

function Install-ADTDeployment
{
    [CmdletBinding()]
    param()

    ##================================================
    ## MARK: Early detection (BEFORE UI)
    ##================================================
    if (Test-AppInstalled) {
        Write-ADTLogEntry -Message "$($Global:Pkg.Name) $($Global:Pkg.Version) already installed (>= $($Global:Pkg.DetectMinVersion)). Exiting." -Severity 1
        return
    }

    ##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    $welcome = @{
        CheckDiskSpace = $true
        PersistPrompt  = $true
    }
    if ($adtSession.AppProcessesToClose.Count -gt 0) {
        $welcome.CloseProcesses = $adtSession.AppProcessesToClose
        $welcome.CloseProcessesCountdown = 60
    }

    Show-ADTInstallationWelcome @welcome
        if ($adtSession.DeployMode -in @('Silent','NonInteractive','Auto')) {
            Stop-RunningApp -ProcessNames $Global:Pkg.ProcessesToClose -GraceSeconds 5
        }

    Show-ADTInstallationProgress

    ##================================================
    ## MARK: Install
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Invoke-InstallPayload

    if (-not (Test-AppInstalled)) {
        throw "Install completed but detection still fails for '$($Global:Pkg.Name)'."
    }

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Installation tasks here>

    ## Display a message at the end of the install.
    if ($adtSession.DeployMode -eq 'Interactive') {
        Show-ADTInstallationPrompt -Message "$($Global:Pkg.Name) installed." -ButtonRightText 'OK'
    }
}

function Uninstall-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    $welcome = @{
        CheckDiskSpace = $false
        PersistPrompt  = $true
    }

    if ($adtSession.AppProcessesToClose -and $adtSession.AppProcessesToClose.Count -gt 0) {
        $welcome.CloseProcesses = $adtSession.AppProcessesToClose
        $welcome.CloseProcessesCountdown = 60
    }

    Show-ADTInstallationWelcome @welcome
        if ($adtSession.DeployMode -in @('Silent','NonInteractive','Auto')) {
            Stop-RunningApp -ProcessNames $Global:Pkg.ProcessesToClose -GraceSeconds 5
        }

    ## Show Progress Message (with the default message).
    Show-ADTInstallationProgress

    ## <Perform Pre-Uninstallation tasks here>


    ##================================================
    ## MARK: Uninstall
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    ## <Perform Uninstallation tasks here>

    Invoke-UninstallPayload
    # Post-uninstall verification (do not fail if it was already absent)
    if (Test-AppInstalled) {
        Write-ADTLogEntry -Message "Post-uninstall verification: '$($Global:Pkg.Name)' still appears installed." -Severity 2
    } else {
        Write-ADTLogEntry -Message "Post-uninstall verification: '$($Global:Pkg.Name)' is not installed." -Severity 1
    }

    Write-ADTLogEntry -Message "Uninstall flow finished. Returning success unless MSI uninstall threw an error." -Severity 1

    ##================================================
    ## MARK: Post-Uninstallation
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Uninstallation tasks here>


    # If it’s not installed anymore, make uninstall idempotent-success and exit clean
    if (-not (Test-AppInstalled)) {
        Write-ADTLogEntry -Message "Uninstall is idempotent-success: app not present. Returning." -Severity 1
        return
    }

}

function Repair-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Repair
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"

    ## Show Progress Message (with the default message).
    Show-ADTInstallationProgress

    ## <Perform Pre-Repair tasks here>


    ##================================================
    ## MARK: Repair
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType

    ## <Perform Repair tasks here>
    ## For most EXE packages, Repair is not supported. You can choose to reinstall:
    Invoke-InstallPayload


    ##================================================
    ## MARK: Post-Repair
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"

    ## <Perform Post-Repair tasks here>
}


##================================================
## MARK: Initialization
##================================================

# Set strict error handling across entire operation.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
$ProgressPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue
Set-StrictMode -Version 1

# Import the module and instantiate a new session.
try
{
    # Import the module locally if available, otherwise try to find it from PSModulePath.
    if (Test-Path -LiteralPath "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1" -PathType Leaf)
    {
        Get-ChildItem -LiteralPath "$PSScriptRoot\PSAppDeployToolkit" -Recurse -File | Unblock-File -ErrorAction Ignore
        Import-Module -FullyQualifiedName @{ ModuleName = "$PSScriptRoot\PSAppDeployToolkit\PSAppDeployToolkit.psd1"; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force
    }
    else
    {
        Import-Module -FullyQualifiedName @{ ModuleName = 'PSAppDeployToolkit'; Guid = '8c3c366b-8606-4576-9f2d-4051144f7ca2'; ModuleVersion = '4.1.8' } -Force
    }

    # Open a new deployment session, replacing $adtSession with a DeploymentSession.
    $iadtParams = Get-ADTBoundParametersAndDefaultValues -Invocation $MyInvocation
    $adtSession = Remove-ADTHashtableNullOrEmptyValues -Hashtable $adtSession
    $adtSession = Open-ADTSession @adtSession @iadtParams -PassThru
}
catch
{
    $Host.UI.WriteErrorLine((Out-String -InputObject $_ -Width ([System.Int32]::MaxValue)))
    exit 60008
}


##================================================
## MARK: Invocation
##================================================

# Commence the actual deployment operation.
try
{
    # Import any found extensions before proceeding with the deployment.
    Get-ChildItem -LiteralPath $PSScriptRoot -Directory | & {
        process
        {
            if ($_.Name -match 'PSAppDeployToolkit\..+$')
            {
                Get-ChildItem -LiteralPath $_.FullName -Recurse -File | Unblock-File -ErrorAction Ignore
                Import-Module -Name $_.FullName -Force
            }
        }
    }

    # Invoke the deployment and close out the session.
    & "$($adtSession.DeploymentType)-ADTDeployment"
    Close-ADTSession
}
catch
{
    # An unhandled error has been caught.
    $mainErrorMessage = "An unhandled error within [$($MyInvocation.MyCommand.Name)] has occurred.`n$(Resolve-ADTErrorRecord -ErrorRecord $_)"
    Write-ADTLogEntry -Message $mainErrorMessage -Severity 3

    ## Error details hidden from the user by default. Show a simple dialog with full stack trace:
    # Show-ADTDialogBox -Text $mainErrorMessage -Icon Stop -NoWait

    ## Or, a themed dialog with basic error message:
    # Show-ADTInstallationPrompt -Message "$($adtSession.DeploymentType) failed at line $($_.InvocationInfo.ScriptLineNumber), char $($_.InvocationInfo.OffsetInLine):`n$($_.InvocationInfo.Line.Trim())`n`nMessage:`n$($_.Exception.Message)" -ButtonRightText OK -Icon Error -NoWait

    Close-ADTSession -ExitCode 60001
}

