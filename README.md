# Windows-Print-Driver-Deep-Cleanup
Uninstall Windows Print driver and related printers 
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
    This switch is used only when the installed PrintManagement module supports
    the RemoveFromDriverStore parameter.

.PARAMETER Force
    Skips the final interactive confirmation. Driver selection is still
    interactive when -DriverName is not supplied.

.EXAMPLE
    .\PrinterDriverDeepCleanup.ps1

.EXAMPLE
    .\PrinterDriverDeepCleanup.ps1 -DriverName "ZDesigner 220Xi4 300 dpi" -PrinterEnvironment "Windows x64"

.EXAMPLE
    .\PrinterDriverDeepCleanup.ps1 -DriverName "Old Driver" -RemoveFromDriverStore -Force

.NOTES
    Run this script in 64-bit Windows PowerShell 5.1 or later as Administrator.
    Removing queues is disruptive. Review the displayed list before confirming.
