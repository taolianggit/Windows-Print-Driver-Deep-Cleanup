#requires -Version 5.1
#requires -RunAsAdministrator

<#
.SYNOPSIS
    Safely removes a selected Windows printer driver and the references that
    commonly cause "The specified driver is in use by one or more printers."

.DESCRIPTION
    The script:
      - Lists installed printer drivers, or accepts -DriverName.
      - Finds and removes queues that use the selected driver.
      - Finds matching Client Side Rendering (CSR) printer objects and ports.
      - Removes matching references from currently loaded HKEY_USERS hives.
      - Exports every registry key before deleting it.
      - Restarts the Print Spooler inside protected try/finally blocks.
      - Runs up to two cleanup passes and verifies that references did not
        return before removing the driver.

    This script intentionally does not load offline NTUSER.DAT files. If an
    unloaded user profile or Group Policy recreates the printer, the final
    verification stops the operation and reports the remaining references.

.PARAMETER DriverName
    Exact printer-driver name. If omitted, an interactive list is displayed.

.PARAMETER PrinterEnvironment
    Optional exact printer environment, for example "Windows x64". Normally
    only needed if more than one installed driver has the same name.

.PARAMETER BackupRoot
    Folder in which registry exports and queue snapshots are written.

.PARAMETER MaxCleanupPasses
    Number of cleanup/detection passes. The default is 2.

.PARAMETER RemoveFromDriverStore
    Also asks Remove-PrinterDriver to remove the package from Driver Store.
    When omitted in interactive mode, prompts whether to remove the package
    (default: No). Explicit -RemoveFromDriverStore:$false skips that prompt.
    Requires the installed PrintManagement module to support this parameter;
    an unsupported explicit request stops before cleanup begins.

.PARAMETER Force
    Skips the Driver Store prompt and final interactive confirmation. Driver
    selection is still interactive when -DriverName is not supplied.
    With -Force, Driver Store removal requires -RemoveFromDriverStore.

.EXAMPLE
    .\PrinterDriverDeepCleanup.ps1

.EXAMPLE
    .\PrinterDriverDeepCleanup.ps1 -DriverName "ZDesigner 220Xi4 300 dpi" -PrinterEnvironment "Windows x64"

.EXAMPLE
    .\PrinterDriverDeepCleanup.ps1 -DriverName "Old Driver" -RemoveFromDriverStore -Force

.NOTES
    Run this script in 64-bit Windows PowerShell 5.1 or later as Administrator.
    Removing queues is disruptive. Review the displayed list before confirming.
#>

[CmdletBinding()]
param(
    [string]$DriverName,

    [string]$PrinterEnvironment,

    [ValidateNotNullOrEmpty()]
    [string]$BackupRoot = 'C:\Temp\PrinterDriverCleanup',

    [ValidateRange(1, 5)]
    [int]$MaxCleanupPasses = 2,

    [switch]$RemoveFromDriverStore,

    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:CSRRoot = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\Servers'
$script:ManifestFile = $null
$BackupFolder = $null

function Write-Section {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    Write-Host ''
    Write-Host ('=' * 72) -ForegroundColor Cyan
    Write-Host $Title -ForegroundColor Cyan
    Write-Host ('=' * 72) -ForegroundColor Cyan
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        $InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function Convert-ToSafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $safeName = $Name -replace '[\\/:*?"<>|]', '_'
    $safeName = $safeName.Trim()

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return 'Registry'
    }

    if ($safeName.Length -gt 50) {
        $safeName = $safeName.Substring(0, 50)
    }

    return $safeName
}

function Get-ShortHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $hashBytes = $sha256.ComputeHash($bytes)
        $hash = [System.BitConverter]::ToString($hashBytes).Replace('-', '')
    }
    finally {
        $sha256.Dispose()
    }

    return $hash.Substring(0, 12)
}

function Convert-ToRegExePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Trailing spaces can be part of a real registry key name.
    $result = $Path
    $result = $result -replace '^Microsoft\.PowerShell\.Core\\Registry::', ''
    $result = $result -replace '^Registry::', ''
    $result = $result -replace '^HKEY_LOCAL_MACHINE', 'HKLM'
    $result = $result -replace '^HKEY_USERS', 'HKU'
    $result = $result -replace '^HKEY_CURRENT_USER', 'HKCU'
    $result = $result -replace '^HKEY_CLASSES_ROOT', 'HKCR'
    $result = $result -replace '^HKEY_CURRENT_CONFIG', 'HKCC'
    $result = $result -replace '^HKLM:', 'HKLM'
    $result = $result -replace '^HKU:', 'HKU'
    $result = $result -replace '^HKCU:', 'HKCU'
    $result = $result -replace '^HKCR:', 'HKCR'
    $result = $result -replace '^HKCC:', 'HKCC'

    if ($result -notmatch '^(HKLM|HKU|HKCU|HKCR|HKCC)(\\|$)') {
        throw "Cannot convert registry provider path for reg.exe: $Path"
    }

    # Registry providers can accept repeated separators that reg.exe rejects.
    $result = $result -replace '\\+', '\'
    return $result.TrimEnd([char]'\')
}

function Invoke-RegistryExport {
    param([string]$RegExe, [string]$RegPath, [string]$File)

    # Windows PowerShell 5.1 turns redirected native stderr into ErrorRecords.
    # Keep the preference local so we can capture the exit code and full message.
    $ErrorActionPreference = 'Continue'
    $PSNativeCommandUseErrorActionPreference = $false
    $output = @(& $RegExe export $RegPath $File /y 2>&1)
    $code = $LASTEXITCODE
    $details = ($output | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
    [PSCustomObject]@{ ExitCode = $code; Details = $details }
}

function Backup-RegistryKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,

        [Parameter(Mandatory = $true)]
        [string]$Category
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Registry key disappeared before backup: $Path"
    }

    $safeCategory = Convert-ToSafeFileName -Name $Category
    $shortHash = Get-ShortHash -Text $Path
    $file = Join-Path -Path $BackupFolder -ChildPath "$safeCategory-$shortHash.reg"
    $regPath = Convert-ToRegExePath -Path $Path

    $regExe = Join-Path -Path $env:SystemRoot -ChildPath 'System32\reg.exe'
    if (-not (Test-Path -LiteralPath $regExe)) {
        $regExe = 'reg.exe'
    }

    Write-Host "Backing up: $regPath" -ForegroundColor DarkCyan
    $requestedRegPath = $regPath
    $result = Invoke-RegistryExport -RegExe $regExe -RegPath $regPath -File $file

    # A broader export is permitted only for the user Connections container.
    # Deletion remains restricted to the originally discovered connection key.
    if ($result.ExitCode -ne 0 -and
        $regPath -match '^(HKU\\[^\\]+\\Printers\\Connections)\\[^\\]+$') {
        $parentRegPath = $Matches[1]
        Write-Warning "Export failed for '$regPath' (exit $($result.ExitCode)): $($result.Details)"
        Write-Host "Retrying backup of parent: $parentRegPath" -ForegroundColor Yellow
        $regPath = $parentRegPath
        $result = Invoke-RegistryExport -RegExe $regExe -RegPath $regPath -File $file
    }

    if ($result.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $file)) {
        throw "Registry backup failed for '$requestedRegPath'; export target '$regPath' (exit code $($result.ExitCode)). $($result.Details) No deletion is allowed without a successful backup."
    }

    # Require the exact requested key in the export, including on parent fallback.
    $fullKey = $requestedRegPath -replace '^HKLM\\', 'HKEY_LOCAL_MACHINE\'
    $fullKey = $fullKey -replace '^HKU\\', 'HKEY_USERS\'
    $fullKey = $fullKey -replace '^HKCU\\', 'HKEY_CURRENT_USER\'
    $fullKey = $fullKey -replace '^HKCR\\', 'HKEY_CLASSES_ROOT\'
    $fullKey = $fullKey -replace '^HKCC\\', 'HKEY_CURRENT_CONFIG\'
    $exportLines = @(Get-Content -LiteralPath $file -Encoding Unicode -ErrorAction Stop)
    if ($exportLines -inotcontains "[$fullKey]") {
        throw "Backup '$file' does not contain the requested key '$fullKey'. Cleanup stopped before deletion."
    }

    [PSCustomObject]@{
        Time       = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Category   = $Category
        RequestedRegistryKey = $requestedRegPath
        RegistryKey = $regPath
        BackupFile = $file
    } | Export-Csv -LiteralPath $script:ManifestFile -NoTypeInformation -Append -Encoding UTF8

    Write-Host "Backed up: $regPath" -ForegroundColor DarkGreen
}

function Get-CSRPrinterObjects {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriverName
    )

    $results = @()

    if (-not (Test-Path -LiteralPath $script:CSRRoot)) {
        return @()
    }

    $serverKeys = @(Get-ChildItem -LiteralPath $script:CSRRoot -ErrorAction Stop)

    foreach ($serverKey in $serverKeys) {
        $printerRoot = Join-Path -Path $serverKey.PSPath -ChildPath 'Printers'
        if (-not (Test-Path -LiteralPath $printerRoot)) {
            continue
        }

        $printerKeys = @(Get-ChildItem -LiteralPath $printerRoot -ErrorAction Stop)

        foreach ($printerKey in $printerKeys) {
            $properties = Get-ItemProperty -LiteralPath $printerKey.PSPath -ErrorAction Stop
            $driverFromPrinter = Get-ObjectPropertyValue -InputObject $properties -Name 'Printer Driver'
            $driverFromDs = $null
            $dsPath = Join-Path -Path $printerKey.PSPath -ChildPath 'DsSpooler'

            if (Test-Path -LiteralPath $dsPath) {
                $dsProperties = Get-ItemProperty -LiteralPath $dsPath -ErrorAction Stop
                $driverFromDs = Get-ObjectPropertyValue -InputObject $dsProperties -Name 'driverName'
            }

            if ($driverFromPrinter -ine $DriverName -and $driverFromDs -ine $DriverName) {
                continue
            }

            $uncLookup = @{}
            foreach ($propertyName in @('Description', 'Name', 'PrinterName', 'Share', 'ShareName', 'Share Name')) {
                $value = Get-ObjectPropertyValue -InputObject $properties -Name $propertyName
                if ($null -eq $value) {
                    continue
                }

                foreach ($item in @($value)) {
                    $text = [string]$item
                    if ($text -match '^\\\\[^\\]+\\[^\\]+$') {
                        $uncLookup[$text] = $true
                    }
                }
            }

            $serverName = $null
            foreach ($propertyName in @('Server', 'ServerName')) {
                $value = Get-ObjectPropertyValue -InputObject $properties -Name $propertyName
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $serverName = [string]$value
                    break
                }
            }

            if ([string]::IsNullOrWhiteSpace($serverName)) {
                $serverName = [string]$serverKey.PSChildName
            }

            $shareName = $null
            foreach ($propertyName in @('Share', 'ShareName', 'Share Name')) {
                $value = Get-ObjectPropertyValue -InputObject $properties -Name $propertyName
                if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
                    $shareName = [string]$value
                    break
                }
            }

            if (-not [string]::IsNullOrWhiteSpace($serverName) -and
                -not [string]::IsNullOrWhiteSpace($shareName)) {
                $cleanServer = $serverName.Trim().TrimStart([char]'\')
                $cleanShare = $shareName.Trim().TrimStart([char]'\')

                if ($cleanServer -notmatch '\\' -and $cleanShare -notmatch '\\') {
                    $uncLookup["\\$cleanServer\$cleanShare"] = $true
                }
            }

            $results += [PSCustomObject]@{
                Server       = [string]$serverKey.PSChildName
                Description  = Get-ObjectPropertyValue -InputObject $properties -Name 'Description'
                DriverName   = $DriverName
                GUID         = [string]$printerKey.PSChildName
                Port         = Get-ObjectPropertyValue -InputObject $properties -Name 'Port'
                RegistryPath = [string]$printerKey.PSPath
                ServerPath   = [string]$serverKey.PSPath
                UNCs         = @($uncLookup.Keys | Sort-Object)
            }
        }
    }

    return @($results | Sort-Object -Property RegistryPath -Unique)
}

function Get-CSRPorts {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$CSRPrinters
    )

    $results = @()

    foreach ($printer in $CSRPrinters) {
        $portRoot = Join-Path -Path $printer.ServerPath -ChildPath 'Client Side Ports'
        if (-not (Test-Path -LiteralPath $portRoot)) {
            continue
        }

        $portKeys = @(Get-ChildItem -LiteralPath $portRoot -Recurse -ErrorAction Stop)
        foreach ($portKey in $portKeys) {
            $guidMatches = $portKey.PSChildName -ieq $printer.GUID
            $portMatches = (
                (-not [string]::IsNullOrWhiteSpace([string]$printer.Port)) -and
                ($portKey.PSChildName -ieq [string]$printer.Port)
            )

            if ($guidMatches -or $portMatches) {
                $results += [PSCustomObject]@{
                    Server = $printer.Server
                    GUID   = $printer.GUID
                    Path   = [string]$portKey.PSPath
                }
            }
        }
    }

    return @($results | Sort-Object -Property Path -Unique)
}

function Get-LoadedUserSIDs {
    $results = @()
    $userKeys = @(Get-ChildItem -LiteralPath 'Registry::HKEY_USERS' -ErrorAction Stop)

    foreach ($userKey in $userKeys) {
        $sid = [string]$userKey.PSChildName
        if (($sid -match '^S-1-5-21-') -or ($sid -match '^S-1-12-1-')) {
            if ($sid -notlike '*_Classes') {
                $results += $sid
            }
        }
    }

    return @($results | Sort-Object -Unique)
}

function Get-UserPrinterReferences {
    param(
        [Parameter(Mandatory = $true)]
        [array]$PrinterUNCs
    )

    $results = @()
    $uncLookup = @{}

    foreach ($unc in $PrinterUNCs) {
        $text = [string]$unc
        if ($text -match '^\\\\[^\\]+\\[^\\]+$') {
            $uncLookup[$text] = $true
        }
    }

    if ($uncLookup.Count -eq 0) {
        return @()
    }

    $valueLocations = @(
        [PSCustomObject]@{ Category = 'ConvertUserDevModesCount'; RelativePath = 'Printers\ConvertUserDevModesCount' },
        [PSCustomObject]@{ Category = 'DevModePerUser'; RelativePath = 'Printers\DevModePerUser' },
        [PSCustomObject]@{ Category = 'DevModes2'; RelativePath = 'Printers\DevModes2' },
        [PSCustomObject]@{ Category = 'Devices'; RelativePath = 'Software\Microsoft\Windows NT\CurrentVersion\Devices' },
        [PSCustomObject]@{ Category = 'PrinterPorts'; RelativePath = 'Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts' }
    )

    foreach ($sid in (Get-LoadedUserSIDs)) {
        foreach ($location in $valueLocations) {
            $path = "Registry::HKEY_USERS\$sid\$($location.RelativePath)"
            if (-not (Test-Path -LiteralPath $path)) {
                continue
            }

            $properties = Get-ItemProperty -LiteralPath $path -ErrorAction Stop
            foreach ($property in $properties.PSObject.Properties) {
                if ($uncLookup.ContainsKey($property.Name)) {
                    $results += [PSCustomObject]@{
                        SID       = $sid
                        Type      = 'RegistryValue'
                        Category  = $location.Category
                        Printer   = $property.Name
                        Path      = $path
                        ValueName = $property.Name
                    }
                }
            }
        }

        $connectionRoot = "Registry::HKEY_USERS\$sid\Printers\Connections"
        if (-not (Test-Path -LiteralPath $connectionRoot)) {
            continue
        }

        # Enumerate real key names instead of constructing a possibly nonexistent
        # path from the UNC. Match padding only for identification; retain the
        # exact key name (including trailing spaces) for backup and deletion.
        $connectionKeys = @(Get-ChildItem -LiteralPath $connectionRoot -ErrorAction Stop)
        foreach ($connectionKey in $connectionKeys) {
            $actualName = [string]$connectionKey.PSChildName
            foreach ($unc in $uncLookup.Keys) {
                if ($unc -notmatch '^\\\\([^\\]+)\\([^\\]+)$') {
                    continue
                }

                $expectedName = ",,$($Matches[1]),$($Matches[2])"
                if ($actualName.TrimEnd([char]' ') -ine $expectedName.TrimEnd([char]' ')) {
                    continue
                }

                # RegistryKey.Name retains the exact native name.
                $connectionPath = 'Registry::' + [string]$connectionKey.Name
                if ($actualName -ine $expectedName) {
                    Write-Host "Matched connection key with trailing spaces: [$actualName]" -ForegroundColor Yellow
                }
                $results += [PSCustomObject]@{
                    SID       = $sid
                    Type      = 'RegistryKey'
                    Category  = 'Connections'
                    Printer   = $unc
                    Path      = $connectionPath
                    ValueName = $null
                }
                break
            }
        }
    }

    return @($results | Sort-Object -Property SID, Path, ValueName -Unique)
}

function Remove-UserPrinterReference {
    param(
        [Parameter(Mandatory = $true)]
        $Reference
    )

    if ($Reference.Type -eq 'RegistryValue') {
        if (-not (Test-Path -LiteralPath $Reference.Path)) {
            return
        }

        $properties = Get-ItemProperty -LiteralPath $Reference.Path -ErrorAction Stop
        if ($null -eq $properties.PSObject.Properties[$Reference.ValueName]) {
            return
        }

        Remove-ItemProperty -LiteralPath $Reference.Path -Name $Reference.ValueName -ErrorAction Stop
        $properties = Get-ItemProperty -LiteralPath $Reference.Path -ErrorAction Stop

        if ($null -ne $properties.PSObject.Properties[$Reference.ValueName]) {
            throw "Registry value still exists after deletion: $($Reference.Path) -> $($Reference.ValueName)"
        }

        return
    }

    if ($Reference.Type -eq 'RegistryKey') {
        # Use native RegistryKey methods so the exact leaf name is preserved.
        $separator = $Reference.Path.LastIndexOf([char]'\')
        $parentPath = $Reference.Path.Substring(0, $separator)
        $leafName = $Reference.Path.Substring($separator + 1)
        if (-not (Test-Path -LiteralPath $parentPath)) { return }
        $nativeParent = Convert-ToRegExePath -Path $parentPath
        if (-not $nativeParent.StartsWith('HKU\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Unexpected user registry root: $nativeParent"
        }
        $parentKey = [Microsoft.Win32.Registry]::Users.OpenSubKey($nativeParent.Substring(4), $true)
        if ($null -eq $parentKey) { return }
        try {
            if (@($parentKey.GetSubKeyNames()) -icontains $leafName) {
                $parentKey.DeleteSubKeyTree($leafName)
            }
            if (@($parentKey.GetSubKeyNames()) -icontains $leafName) {
                throw "Registry key still exists after deletion: $($Reference.Path)"
            }
        }
        finally {
            $parentKey.Close()
        }
        return
    }

    throw "Unknown user-reference type: $($Reference.Type)"
}

function Wait-ServiceState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Running', 'Stopped')]
        [string]$DesiredState,

        [int]$TimeoutSeconds = 30
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $service = Get-Service -Name $Name -ErrorAction Stop
        if ($service.Status.ToString() -eq $DesiredState) {
            return
        }

        Start-Sleep -Milliseconds 250
    } while ((Get-Date) -lt $deadline)

    throw "Service '$Name' did not reach state '$DesiredState' within $TimeoutSeconds seconds."
}

function Ensure-SpoolerRunning {
    $service = Get-Service -Name Spooler -ErrorAction Stop
    if ($service.Status -ne 'Running') {
        Start-Service -Name Spooler -ErrorAction Stop
        Wait-ServiceState -Name Spooler -DesiredState Running
    }
}

function Remove-PrinterQueueStrict {
    param(
        [Parameter(Mandatory = $true)]
        $Printer
    )

    $name = [string]$Printer.Name
    $firstError = $null

    try {
        Remove-Printer -InputObject $Printer -Confirm:$false -ErrorAction Stop
    }
    catch {
        $firstError = $_.Exception.Message
        $mode = if ($name -like '\\*') { '/dn' } else { '/dl' }
        $rundll32 = Join-Path -Path $env:SystemRoot -ChildPath 'System32\rundll32.exe'

        & $rundll32 printui.dll,PrintUIEntry $mode /n $name 2>&1 | Out-Null
    }

    Start-Sleep -Milliseconds 500

    $remaining = @(
        Get-Printer -ErrorAction Stop |
            Where-Object { $_.Name -ieq $name }
    )

    if ($remaining.Count -gt 0) {
        if ([string]::IsNullOrWhiteSpace($firstError)) {
            $firstError = 'The queue remained after Remove-Printer returned.'
        }

        throw "Failed to remove printer queue '$name'. $firstError"
    }

    Write-Host "Removed queue: $name" -ForegroundColor Green
}

function Invoke-SpoolerRegistryCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$CSRPrinters,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$CSRPorts
    )

    $paths = @()
    foreach ($printer in $CSRPrinters) {
        $paths += [string]$printer.RegistryPath
    }
    foreach ($port in $CSRPorts) {
        $paths += [string]$port.Path
    }
    $paths = @($paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

    $removalErrors = @()

    try {
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Wait-ServiceState -Name Spooler -DesiredState Stopped

        foreach ($path in $paths) {
            if (-not (Test-Path -LiteralPath $path)) {
                continue
            }

            try {
                Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                if (Test-Path -LiteralPath $path) {
                    throw 'The key still exists after Remove-Item returned.'
                }
                Write-Host "Removed CSR key: $path" -ForegroundColor Green
            }
            catch {
                $removalErrors += "$path : $($_.Exception.Message)"
            }
        }
    }
    finally {
        Ensure-SpoolerRunning
    }

    if ($removalErrors.Count -gt 0) {
        throw "One or more CSR keys could not be removed:`n$($removalErrors -join [Environment]::NewLine)"
    }
}

function Get-CleanupState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriverName,

        [array]$KnownUNCs = @()
    )

    $activePrinters = @(
        Get-Printer -ErrorAction Stop |
            Where-Object { $_.DriverName -ieq $DriverName }
    )

    $csrPrinters = @(Get-CSRPrinterObjects -DriverName $DriverName)
    $csrPorts = @(Get-CSRPorts -CSRPrinters $csrPrinters)

    $printerUNCs = @($KnownUNCs)
    foreach ($printer in $activePrinters) {
        if ([string]$printer.Name -match '^\\\\[^\\]+\\[^\\]+$') {
            $printerUNCs += [string]$printer.Name
        }
    }
    foreach ($csrPrinter in $csrPrinters) {
        $printerUNCs += @($csrPrinter.UNCs)
    }

    $printerUNCs = @(
        $printerUNCs |
            Where-Object { [string]$_ -match '^\\\\[^\\]+\\[^\\]+$' } |
            Sort-Object -Unique
    )

    $userReferences = @()
    if ($printerUNCs.Count -gt 0) {
        $userReferences = @(Get-UserPrinterReferences -PrinterUNCs $printerUNCs)
    }

    return [PSCustomObject]@{
        ActivePrinters = @($activePrinters)
        CSRPrinters    = @($csrPrinters)
        CSRPorts       = @($csrPorts)
        UserReferences = @($userReferences)
        PrinterUNCs    = @($printerUNCs)
    }
}

function Backup-CleanupState {
    param(
        [Parameter(Mandatory = $true)]
        $State,

        [Parameter(Mandatory = $true)]
        [int]$Pass,

        [Parameter(Mandatory = $true)]
        [string]$BackupFolder
    )

    if ($State.ActivePrinters.Count -gt 0) {
        $queueFile = Join-Path -Path $BackupFolder -ChildPath "Pass-$Pass-PrinterQueues.csv"
        $State.ActivePrinters |
            Select-Object Name, DriverName, PortName, Shared, ShareName, Type |
            Export-Csv -LiteralPath $queueFile -NoTypeInformation -Encoding UTF8
    }

    if ($State.CSRPrinters.Count -gt 0) {
        $csrFile = Join-Path -Path $BackupFolder -ChildPath "Pass-$Pass-CSRObjects.csv"
        $State.CSRPrinters |
            Select-Object Server, Description, DriverName, GUID, Port,
                @{ Name = 'UNCs'; Expression = { $_.UNCs -join ';' } }, RegistryPath |
            Export-Csv -LiteralPath $csrFile -NoTypeInformation -Encoding UTF8
    }

    $registryItems = @()
    $localQueueRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers'
    foreach ($item in $State.ActivePrinters) {
        if ([string]$item.Name -like '\\*') {
            continue
        }

        $queueRegistryPath = Join-Path -Path $localQueueRoot -ChildPath ([string]$item.Name)
        if (Test-Path -LiteralPath $queueRegistryPath) {
            $registryItems += [PSCustomObject]@{
                Category = "Pass-$Pass-Local-Queue"
                Path     = $queueRegistryPath
            }
        }
    }
    foreach ($item in $State.CSRPrinters) {
        $registryItems += [PSCustomObject]@{
            Category = "Pass-$Pass-CSR"
            Path     = [string]$item.RegistryPath
        }
    }
    foreach ($item in $State.CSRPorts) {
        $registryItems += [PSCustomObject]@{
            Category = "Pass-$Pass-CSR-Port"
            Path     = [string]$item.Path
        }
    }
    foreach ($item in $State.UserReferences) {
        $registryItems += [PSCustomObject]@{
            Category = "Pass-$Pass-User-$($item.Category)"
            Path     = [string]$item.Path
        }
    }

    $uniqueItems = @{}
    foreach ($item in $registryItems) {
        if ([string]::IsNullOrWhiteSpace($item.Path)) {
            continue
        }

        $key = $item.Path.ToLowerInvariant()
        if (-not $uniqueItems.ContainsKey($key)) {
            $uniqueItems[$key] = $item
        }
    }

    foreach ($item in $uniqueItems.Values) {
        Backup-RegistryKey -Path $item.Path -BackupFolder $BackupFolder -Category $item.Category
    }
}

function Get-ExactDriverMatches {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string]$Environment
    )

    $matches = @(
        Get-PrinterDriver -ErrorAction Stop |
            Where-Object { $_.Name -ieq $Name }
    )

    if (-not [string]::IsNullOrWhiteSpace($Environment)) {
        $matches = @(
            $matches |
                Where-Object {
                    (Get-ObjectPropertyValue -InputObject $_ -Name 'PrinterEnvironment') -ieq $Environment
                }
        )
    }

    return @($matches)
}

try {
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        throw 'Run the script from 64-bit Windows PowerShell. The 32-bit registry view can miss printer data.'
    }

    Import-Module PrintManagement -ErrorAction Stop
    Ensure-SpoolerRunning

    $drivers = @(Get-PrinterDriver -ErrorAction Stop | Sort-Object Name, PrinterEnvironment)
    if ($drivers.Count -eq 0) {
        throw 'No installed printer drivers were found.'
    }

    $selectedDriver = $null

    if (-not [string]::IsNullOrWhiteSpace($DriverName)) {
        $driverMatches = @(Get-ExactDriverMatches -Name $DriverName -Environment $PrinterEnvironment)
        if ($driverMatches.Count -eq 0) {
            throw "Printer driver not found: $DriverName"
        }
        if ($driverMatches.Count -gt 1) {
            throw "More than one driver is named '$DriverName'. Specify -PrinterEnvironment as well."
        }
        $selectedDriver = $driverMatches[0]
    }
    else {
        Write-Section 'Available Printer Drivers'
        for ($index = 0; $index -lt $drivers.Count; $index++) {
            $driver = $drivers[$index]
            $environment = Get-ObjectPropertyValue -InputObject $driver -Name 'PrinterEnvironment'
            $provider = Get-ObjectPropertyValue -InputObject $driver -Name 'Provider'
            Write-Host ('[{0}] {1}  [{2}; {3}]' -f ($index + 1), $driver.Name, $environment, $provider)
        }

        while ($null -eq $selectedDriver) {
            Write-Host ''
            $selectionText = Read-Host 'Select the driver number to remove'
            $selection = 0

            if ([int]::TryParse($selectionText, [ref]$selection) -and
                $selection -ge 1 -and $selection -le $drivers.Count) {
                $selectedDriver = $drivers[$selection - 1]
            }
            else {
                Write-Host 'Invalid selection.' -ForegroundColor Yellow
            }
        }
    }

    $DriverName = [string]$selectedDriver.Name
    $selectedEnvironment = [string](Get-ObjectPropertyValue -InputObject $selectedDriver -Name 'PrinterEnvironment')

    Write-Section 'Selected Driver'
    Write-Host "Name         : $DriverName" -ForegroundColor Yellow
    Write-Host "Environment  : $selectedEnvironment"
    Write-Host "Provider     : $(Get-ObjectPropertyValue -InputObject $selectedDriver -Name 'Provider')"
    Write-Host "Manufacturer : $(Get-ObjectPropertyValue -InputObject $selectedDriver -Name 'Manufacturer')"
    Write-Host "Version      : $(Get-ObjectPropertyValue -InputObject $selectedDriver -Name 'DriverVersion')"
    Write-Host "INF path     : $(Get-ObjectPropertyValue -InputObject $selectedDriver -Name 'InfPath')"
    Write-Host "Config DLL   : $(Get-ObjectPropertyValue -InputObject $selectedDriver -Name 'ConfigFile')"

    # Decide package-removal behavior and validate support before any cleanup.
    $removeCommand = Get-Command -Name Remove-PrinterDriver -ErrorAction Stop
    $supportsDriverStoreRemoval = $removeCommand.Parameters.ContainsKey('RemoveFromDriverStore')
    if ($RemoveFromDriverStore -and -not $supportsDriverStoreRemoval) {
        throw 'This version of PrintManagement does not support -RemoveFromDriverStore. No cleanup has been performed.'
    }

    if (-not $Force -and -not $PSBoundParameters.ContainsKey('RemoveFromDriverStore')) {
        if ($supportsDriverStoreRemoval) {
            Write-Host ''
            Write-Host 'Driver Store removal also asks Windows to delete the selected driver package.'
            Write-Host 'Windows may refuse package removal if it is still in use.'
            while ($true) {
                $storeChoice = (Read-Host 'Also remove the driver package from Driver Store? [Y/N] (default: N)').Trim()
                if ([string]::IsNullOrWhiteSpace($storeChoice) -or $storeChoice -imatch '^(N|No)$') {
                    $RemoveFromDriverStore = $false
                    break
                }
                if ($storeChoice -imatch '^(Y|Yes)$') {
                    $RemoveFromDriverStore = $true
                    break
                }
                Write-Host 'Enter Y or N.' -ForegroundColor Yellow
            }
        }
        else {
            Write-Warning 'Driver Store removal is unavailable in this PrintManagement version. Only printer-driver removal will be requested.'
        }
    }

    $initialState = Get-CleanupState -DriverName $DriverName

    Write-Section 'Discovered References'
    Write-Host "Active printer queues        : $($initialState.ActivePrinters.Count)"
    Write-Host "CSR printer objects          : $($initialState.CSRPrinters.Count)"
    Write-Host "CSR client-side port keys    : $($initialState.CSRPorts.Count)"
    Write-Host "Loaded-user registry entries : $($initialState.UserReferences.Count)"

    if ($initialState.ActivePrinters.Count -gt 0) {
        Write-Host ''
        $initialState.ActivePrinters |
            Select-Object Name, DriverName, PortName, Shared, ShareName |
            Format-Table -AutoSize |
            Out-Host
    }

    if ($initialState.CSRPrinters.Count -gt 0) {
        Write-Host ''
        $initialState.CSRPrinters |
            Select-Object Server, Description, GUID, RegistryPath |
            Format-Table -AutoSize |
            Out-Host
    }

    if ($initialState.UserReferences.Count -gt 0) {
        Write-Host ''
        $initialState.UserReferences |
            Select-Object SID, Category, Printer, Type |
            Format-Table -AutoSize |
            Out-Host
    }

    Write-Host ''
    $storeAction = if ($RemoveFromDriverStore) { 'YES - request package removal' } else { 'NO - keep package in Driver Store' }
    Write-Host "Driver Store cleanup: $storeAction" -ForegroundColor Yellow
    Write-Warning 'All listed queues and references will be removed before the driver is removed.'
    Write-Warning 'Only user profiles currently loaded under HKEY_USERS are inspected.'
    Write-Warning 'A GPO or logon process may recreate a printer; final verification will detect this.'

    if (-not $Force) {
        Write-Host ''
        $confirmation = Read-Host 'Type REMOVE to continue'
        if ($confirmation -ine 'REMOVE') {
            Write-Host 'Operation cancelled.' -ForegroundColor Yellow
            return
        }
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $safeDriverName = Convert-ToSafeFileName -Name $DriverName
    $BackupFolder = Join-Path -Path $BackupRoot -ChildPath "$timestamp-$safeDriverName"
    New-Item -ItemType Directory -Path $BackupFolder -Force -ErrorAction Stop | Out-Null

    $script:ManifestFile = Join-Path -Path $BackupFolder -ChildPath 'RegistryBackupManifest.csv'
    $driverSnapshot = Join-Path -Path $BackupFolder -ChildPath 'SelectedDriver.xml'
    $selectedDriver | Export-Clixml -LiteralPath $driverSnapshot -Encoding UTF8

    Write-Host ''
    Write-Host "Backup folder: $BackupFolder" -ForegroundColor Cyan

    $knownUNCs = @($initialState.PrinterUNCs)
    $spoolerRestarted = $false

    for ($pass = 1; $pass -le $MaxCleanupPasses; $pass++) {
        Write-Section "Cleanup Pass $pass of $MaxCleanupPasses"
        $state = Get-CleanupState -DriverName $DriverName -KnownUNCs $knownUNCs
        $knownUNCs = @($state.PrinterUNCs)

        $referenceCount = (
            $state.ActivePrinters.Count +
            $state.CSRPrinters.Count +
            $state.CSRPorts.Count +
            $state.UserReferences.Count
        )

        Write-Host "Active queues   : $($state.ActivePrinters.Count)"
        Write-Host "CSR objects     : $($state.CSRPrinters.Count)"
        Write-Host "CSR ports       : $($state.CSRPorts.Count)"
        Write-Host "User references : $($state.UserReferences.Count)"

        if ($referenceCount -eq 0) {
            Write-Host 'No remaining printer references were found.' -ForegroundColor Green
            break
        }

        Write-Host ''
        Write-Host 'Backing up registry keys before deletion...' -ForegroundColor Cyan
        Backup-CleanupState -State $state -Pass $pass -BackupFolder $BackupFolder

        foreach ($printer in $state.ActivePrinters) {
            Remove-PrinterQueueStrict -Printer $printer
        }

        foreach ($reference in $state.UserReferences) {
            Remove-UserPrinterReference -Reference $reference
            Write-Host "Removed user reference: $($reference.SID) / $($reference.Category) / $($reference.Printer)" -ForegroundColor Green
        }

        Invoke-SpoolerRegistryCleanup -CSRPrinters $state.CSRPrinters -CSRPorts $state.CSRPorts
        $spoolerRestarted = $true
        Start-Sleep -Seconds 3
    }

    if (-not $spoolerRestarted) {
        Write-Section 'Restarting Print Spooler'
        Invoke-SpoolerRegistryCleanup -CSRPrinters @() -CSRPorts @()
        Start-Sleep -Seconds 3
    }

    Write-Section 'Final Pre-Removal Verification'
    $finalState = Get-CleanupState -DriverName $DriverName -KnownUNCs $knownUNCs

    Write-Host "Active queues   : $($finalState.ActivePrinters.Count)"
    Write-Host "CSR objects     : $($finalState.CSRPrinters.Count)"
    Write-Host "CSR ports       : $($finalState.CSRPorts.Count)"
    Write-Host "User references : $($finalState.UserReferences.Count)"

    $remainingReferenceCount = (
        $finalState.ActivePrinters.Count +
        $finalState.CSRPrinters.Count +
        $finalState.CSRPorts.Count +
        $finalState.UserReferences.Count
    )

    if ($remainingReferenceCount -gt 0) {
        if ($finalState.ActivePrinters.Count -gt 0) {
            $finalState.ActivePrinters |
                Select-Object Name, DriverName, PortName |
                Format-Table -AutoSize |
                Out-Host
        }
        if ($finalState.CSRPrinters.Count -gt 0) {
            $finalState.CSRPrinters |
                Select-Object Server, Description, GUID, RegistryPath |
                Format-Table -AutoSize |
                Out-Host
        }
        if ($finalState.UserReferences.Count -gt 0) {
            $finalState.UserReferences |
                Select-Object SID, Category, Printer, Type |
                Format-Table -AutoSize |
                Out-Host
        }

        throw 'Printer references returned after cleanup. Check printer-deployment GPOs, logon scripts, and unloaded user profiles before trying again.'
    }

    Write-Section 'Removing Printer Driver'
    $removeParameters = @{
        Name               = $DriverName
        PrinterEnvironment = $selectedEnvironment
        Confirm            = $false
        ErrorAction        = 'Stop'
    }

    if ([string]::IsNullOrWhiteSpace($selectedEnvironment)) {
        $removeParameters.Remove('PrinterEnvironment') | Out-Null
    }

    if ($RemoveFromDriverStore) {
        $removeParameters['RemoveFromDriverStore'] = $true
    }

    Remove-PrinterDriver @removeParameters

    $driverStillExists = @(Get-ExactDriverMatches -Name $DriverName -Environment $selectedEnvironment)
    if ($driverStillExists.Count -gt 0) {
        throw "Remove-PrinterDriver returned, but the driver still exists: $DriverName"
    }

    Write-Host ''
    Write-Host 'SUCCESS: Printer driver removed.' -ForegroundColor Green
    Write-Host "Driver : $DriverName" -ForegroundColor Green
    if ($RemoveFromDriverStore) {
        Write-Host 'Driver Store package removal was requested from Windows; package absence was not independently verified.' -ForegroundColor Yellow
    }
    else {
        Write-Host 'Driver Store package removal was not requested.'
    }
    Write-Host "Backup : $BackupFolder" -ForegroundColor Green
}
catch {
    $originalError = $_

    try {
        Ensure-SpoolerRunning
    }
    catch {
        Write-Host 'CRITICAL: The Print Spooler could not be restarted.' -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    Write-Host ''
    Write-Host 'FAILED: Printer-driver cleanup did not complete.' -ForegroundColor Red
    Write-Host $originalError.Exception.Message -ForegroundColor Red
    Write-Host $originalError.InvocationInfo.PositionMessage -ForegroundColor DarkYellow

    if (-not [string]::IsNullOrWhiteSpace($BackupFolder)) {
        Write-Host "Available backups: $BackupFolder" -ForegroundColor Yellow
    }

    exit 1
}
