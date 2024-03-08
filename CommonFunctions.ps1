<#
Dot source this file to use it's functions in a script.  Here is an example of my jumpstart / bootstrap section of my scripts:

## Bootstrap execution

#region Initial 'bootstrapping' of my scripts - loads a Common Functions module that most of my scripts use
$commonFunctions = "..\Script Assets\CommonFunctions.ps1"
try {
    if (Test-Path $commonFunctions) {
        # Dot source the file for use of functions
        . $commonFunctions
        Write-Debug "'Common Functions' loaded successfully."
    } else {
        throw "Failed to load 'Common Functions' from $commonFunctions."
    }
} catch { 
    # Show error message and write error
    [System.Windows.Forms.MessageBox]::Show($Error[0], "File Not Found", "OK", "Error")
    Write-Error "$Error[0]" -ErrorId M0 -Category ResourceUnavailable
    return $false
}

#region This section requires no user interaction
# Variables
$assetsFolder = '..\Script Assets\'
$moduleList = @{
    "Get-ComputerNames" = @{
        Name = "Get-ComputerNames"
        Path = $assetsFolder
    }
    "Active Directory" = @{
        Name = "ActiveDirectory"
        ErrorMessage = "This script requires the Active Directory module.`nYou need to install RSAT tools to proceed.`nRSAT Tools can be found at:`n\\lmh-sms\applications$\Microsoft\RSAT Tools"
    }
    "Progress Bar" = @{
        Name = "ProgressBar"
        Path = $assetsFolder
    }
    "Easy Async Jobs" = @{
        Name = "EasyAsyncJobs"
        Path = $assetsFolder
    }
}

# Start tracking runtime
$startTime = Measure-ScriptRuntime
Write-Debug "Time Logged"

# Verify the asset folder is available and checks that the Building and Department Codes.csv file is present - you can specify a different folder with the -assetsFolder flag
$assetLoaded = Verify-AssetsFolder -assetsFolder $assetsFolder
if (!$assetLoaded) { return }
Write-Debug "Asset folder verified"

# Ensure the user has admin privileges
$adminVerified = Ensure-AdminPrivileges $myinvocation.MyCommand.Definition
if (!$adminVerified) { return }
Write-Debug "Administrator Privileges Verified"

# Load modules listed in the ModuleList variable and track the ErrorIdCounter (no reason to track it or have it at this point)
try { $moduleIdCount = Load-ModuleGeneric -ModuleList $moduleList } catch { return $Error[0] }
if (!$moduleIdCount) { return }
Write-Debug "Modules Loaded"
#endregion

Write-Debug "Bootstrap complete"
#endregion Bootstrap execution

#>

Add-Type -AssemblyName System.Windows.Forms  # Load Windows Forms

#region This area is for showing a form based on an array of hashtables
# Define a custom sorter that implements the IComparer interface for sorting ListView items
class ListViewCustomSorter : System.Collections.IComparer {
    <#
    .SYNOPSIS
    Custom sorter for ListView control that implements the IComparer interface.
    
    .DESCRIPTION
    This class provides custom sorting functionality for a ListView control in a Windows Forms application.
    It supports sorting items based on the text of a specified column in ascending or descending order.
    
    .NOTES
    Author: Your Name
    Version: 1.0
    #>

    [int] $ColumnToSort = 0
    [System.Collections.CaseInsensitiveComparer] $ObjectComparer
    [System.Windows.Forms.SortOrder] $Order = [System.Windows.Forms.SortOrder]::Ascending

    ListViewCustomSorter() {
        $this.ObjectComparer = New-Object System.Collections.CaseInsensitiveComparer
    }

    [int] Compare([object] $x, [object] $y) {
        $itemX = [System.Windows.Forms.ListViewItem]$x
        $itemY = [System.Windows.Forms.ListViewItem]$y

        $compareResult = $this.ObjectComparer.Compare($itemX.SubItems[$this.ColumnToSort].Text, $itemY.SubItems[$this.ColumnToSort].Text)

        if ($this.Order -eq [System.Windows.Forms.SortOrder]::Descending) {
            $compareResult = -$compareResult
        }

        return $compareResult
    }
}

function Show-DataInForm {
    <#
    .SYNOPSIS
    Displays data from a collection of hashtables in a Windows Form with sortable columns.
    
    .DESCRIPTION
    This function creates a Windows Form displaying data from an array of hashtables in a ListView control.
    The columns can be dynamically generated based on the keys of the hashtables or specified by the user.
    Clicking on a column header sorts the items based on that column's values.
    
    .PARAMETER Data
    An array of hashtables where each hashtable represents a row of data to be displayed.
    
    .PARAMETER ColumnHeaders
    An optional array of strings specifying which columns to display and in what order. If not provided, the columns are determined based on the keys in the first hashtable.
    
    .VERSION
    1.0.0 - Initial release.

    .AUTHOR
    Christopher Thompson - christopher.thompson@leehealth.org

    .DEPENDENCIES
    None
    
    .EXAMPLE
    $Data = @(
        @{ ComputerName = "PC1"; Status = "Online"; IP = "192.168.1.1" },
        @{ ComputerName = "PC2"; Status = "Offline"; IP = "192.168.1.2" }
    )
    Show-DataInForm -Data $Data

    .EXAMPLE
    $data = @(
        @{ Computer = "PC 1"; Status = "Online"; Latency = "93ms"; Devices = "4"; Monitors = "3 Monitor(s)"; IP = "192.168.1.118" }
        @{ Computer = "PC 2"; Status = "Unknown"; Latency = "85ms"; Devices = "3"; Monitors = "2 Monitor(s)"; IP = "192.168.1.94" }
        @{ Computer = "PC 3"; Status = "Unknown"; Latency = "37ms"; Devices = "2"; Monitors = "2 Monitor(s)"; IP = "192.168.1.227" }
        @{ Computer = "PC 4"; Status = "Online"; Latency = "9ms"; Devices = "3"; Monitors = "3 Monitor(s)"; IP = "192.168.1.238" }
        @{ Computer = "PC 5"; Status = "Offline"; Latency = "29ms"; Devices = "4"; Monitors = "3 Monitor(s)"; IP = "192.168.1.114" }
    )
    
    $Headers = @(Computer, Status, Latency, Devices, Monitors, IP Address)
    
    Show-DataInForm -Data $data -ColumnHeaders $headers

    .NOTES

    #>
    
    param (
        [Parameter(Mandatory = $true)]
        [Hashtable[]]$Data,  # Array of hashtables, each representing a row of data.

        [Parameter(Mandatory = $false)]
        [string[]]$ColumnHeaders  # Optional. Specifies which columns to display and their order.
    )

    $customSorter = [ListViewCustomSorter]::new()

    # Create the form for displaying data
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Data Display'
    $form.Size = New-Object System.Drawing.Size(800, 600)

    # Initialize the ListView control for tabular data display
    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = [System.Windows.Forms.View]::Details
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Size = New-Object System.Drawing.Size(760, 520)
    $listView.Location = New-Object System.Drawing.Point(10, 10)
    $listView.ListViewItemSorter = $customSorter

    # Determine and add column headers to the ListView
    $columnsToAdd = if ($ColumnHeaders -and $ColumnHeaders.Count -gt 0) {
        $ColumnHeaders
    } else {
        $keys = $Data[0].Keys
        $computerKey = if ($keys -contains "ComputerName") { "ComputerName" } elseif ($keys -contains "Computer") { "Computer" }
        @($computerKey) + $keys.Where({ $_ -ne $computerKey })
    }

    foreach ($header in $columnsToAdd) {
        $listView.Columns.Add($header) | Out-Null
    }

    # Populate the ListView with data from the hashtables
    foreach ($item in $Data) {
        $listViewItem = New-Object System.Windows.Forms.ListViewItem($item[$columnsToAdd[0]].ToString())
        foreach ($header in $columnsToAdd[1..$columnsToAdd.Count]) {
            $subItemText = $item[$header] -as [string]
            $listViewItem.SubItems.Add($subItemText) | Out-Null
        }
        $listView.Items.Add($listViewItem) | Out-Null
    }

    # Automatically adjust column widths to fit their content
    foreach ($column in $listView.Columns) {
        $column.Width = -2
    }

    # Handle column clicks to sort data
    $listView.Add_ColumnClick({
        param($sender, $e)
        $customSorter.ColumnToSort = $e.Column

        # Toggle the sorting order if the same column is clicked again
        if ($customSorter.Order -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $customSorter.Order = [System.Windows.Forms.SortOrder]::Descending
        } else {
            $customSorter.Order = [System.Windows.Forms.SortOrder]::Ascending
        }

        $sender.Sort()
    })

    # Display the form as a modal dialog
    $form.Controls.Add($listView)
    $form.ShowDialog() | Out-Null
}
#endregion

# Function to measure the script runtime
function Measure-ScriptRuntime {
    <#
    .SYNOPSIS
    Measures the runtime of a script.
    
    .DESCRIPTION
    This function calculates the elapsed time since the script started running. It can also mark the start time of a script.
    
    .PARAMETER startTime
    The start time of the script. If not provided, the current time is returned as the start time.
    
    .EXAMPLE
    # To mark the start time:
    $startTime = Measure-ScriptRuntime
    
    # At the end of the script to display runtime:
    Measure-ScriptRuntime -startTime $startTime
    
    .CHANGELOG
    1.0 Initial release.
    
    .DEPENDENCIES
    None.
    
    .VERSION
    1.0
    
    .AUTHOR
    Chris Thompson - christopher.thompson@leehealth.org
    
    .INSTRUCTIONS
    Call this function at the start and end of your script to measure its runtime.
    
    .EXCEPTIONS
    None.
    
    .ISSUES
    None.
    
    .TODO
    None.
    
    .NOTES
    This function is useful for performance monitoring and optimization.
    #>

    param (
        [DateTime]$startTime
    )

    if ($startTime) {
        # Calculate and display the elapsed time
        $endTime = Get-Date
        $elapsedTime = $endTime - $startTime

        $elapsedTimeSpan = [TimeSpan]::FromMilliseconds($elapsedTime.TotalMilliseconds)
        $formattedElapsedTime = $elapsedTimeSpan.ToString("hh\:mm\:ss")

        Write-Host "Script finished at: $endTime"
        Write-Host "Total script runtime: $formattedElapsedTime"
    } else {
        # If no start time is provided, return the current time
        $startTime = Get-Date
        Write-Host "Script started at: $startTime"
        return $startTime
    }
}

# Function to verify Assets folder location and set directory location for script execution
function Verify-AssetsFolder {
    <#
    .SYNOPSIS
    Verifies the presence of an Assets folder and a specific CSV file within it.
    
    .DESCRIPTION
    Checks if the specified Assets folder exists and contains the "Building and Department Codes.csv" file. Prompts the user to select the folder if not found.
    
    .PARAMETER assetsFolder
    The path to the Assets folder. Default is "..\Script Assets\".
    
    .EXAMPLE
    # Verify the default Assets folder:
    Verify-AssetsFolder
    
    # Specify a custom Assets folder path:
    Verify-AssetsFolder -assetsFolder "C:\Custom\Path\To\Assets"
    
    .CHANGELOG
    1.0 Initial release.
    
    .DEPENDENCIES
    Requires Windows Forms for folder selection dialog.
    
    .VERSION
    1.0
    
    .AUTHOR
    Chris Thompson - christopher.thompson@leehealth.org
    
    .INSTRUCTIONS
    Call this function to ensure the required Assets folder and CSV file are present before proceeding with script operations that depend on these resources.
    
    .EXCEPTIONS
    Exits script if Assets folder or CSV file not found.
    
    .ISSUES
    None.
    
    .TODO
    Improve error handling for non-interactive environments.
    
    .NOTES
    Adjust the default assetsFolder path as needed for your environment.
    
    #>

    param(
        [string]$assetsFolder = "..\Script Assets\"  # Default path
    )
    
    if (!(Test-Path $assetsFolder)) {
        # Prompt user to choose path if default path doesn't exist
        $assetsFolder = (New-Object System.Windows.Forms.FolderBrowserDialog).ShowDialog()
        if (!$assetsFolder) {
            Write-Error "Assets folder not selected. Exiting script."
            [System.Windows.Forms.MessageBox]::Show("Assets folder not selected. Please run the script again and choose the correct path.", "Error", 0)
            return $false
        }
    }
    
    $testFilePath = Join-Path $assetsFolder "Building and Department Codes.csv"
    if (!(Test-Path $testFilePath)) {
        Write-Error "Building and Department Codes.csv file not found in the Assets folder."
        [System.Windows.Forms.MessageBox]::Show("Building and Department Codes.csv file is missing. Please ensure it's in the Assets folder.", "Error", 0)
        return $false
    }
    return $true
}

# Function to check for Admin, prompt for creds if not admin, and retry
function Ensure-AdminPrivileges {
    <#
    .SYNOPSIS
    Ensures the script is running with administrative privileges.
    
    .DESCRIPTION
    Checks if the current PowerShell session has administrative privileges. If not, attempts to restart the script with elevated privileges.
    
    .PARAMETER commandRan
    The command or script file to be executed with administrative privileges.
    
    .EXAMPLE
    Ensure-AdminPrivileges -commandRan $MyInvocation.MyCommand.Definition
    
    .CHANGELOG
    1.0 Initial release.
    
    .DEPENDENCIES
    None.
    
    .VERSION
    1.0
    
    .AUTHOR
    Chris Thompson - christopher.thompson@leehealth.org
    
    .INSTRUCTIONS
    Include this function call at the beginning of scripts that require administrative privileges.
    
    .EXCEPTIONS
    Catches exceptions related to starting a process with elevated privileges.
    
    .ISSUES
    None.
    
    .TODO
    Enhance user feedback for elevation prompt cancellation.
    
    .NOTES
    Useful for scripts that need to perform tasks requiring admin rights.
    
    #>

    param ($commandRan)
    # Checks if the script is running with administrative privileges
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdministrator = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (!$isAdministrator) {
        try {
            # Restarts the script with elevated privileges, prompting for credentials
            Start-Process powershell.exe -Verb RunAs -ArgumentList ('-NoProfile -NoExit -File "{0}" -Elevated' -f ($commandRan))
            return $false
        } catch {
            # Handles errors during privilege elevation
            $errorMessage = $_.Exception.Message
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error")
            Write-Error $errorMessage -ErrorId B1 -Category PermissionDenied
            return $false
        }
    }
    return $true
}

# Function for selecting a file
function Get-FileName {
    <#
    .SYNOPSIS
    Prompts the user to select a file through a file dialog window.
    
    .DESCRIPTION
    Opens a file dialog window asking the user to select a file. Validates the selected file exists and returns its full path. If the user cancels the operation or an invalid file is selected, the function returns $false.
    
    .PARAMETER title
    Optional. Specifies the title of the file dialog window. Default is "Select a CSV file".
    
    .PARAMETER filter
    Optional. Defines the filter for displayed file types in the dialog. Default is set to show CSV files and all files.
    
    .OUTPUTS
    String or Boolean. Returns the full path to the selected file as a string if successful; otherwise, returns $false.
    
    .EXAMPLE
    $csvFilePath = Get-FileName -title "Select a PowerShell file" -filter "PS1 files (*.ps1)|*.ps1|Command Script Files (*.cmd)|*.CMD|All files (*.*)|*.*"
    This example shows how to use Get-FileName to prompt the user for a PS1 file or a CMD file, using a custom title and file filter.
    
    .EXAMPLE
    $filePath = Get-FileName
    This example uses the default title and filter to prompt the user to select any file, primarily focusing on CSV files.
    
    .CHANGELOG
    1.0 Initial release.
    
    .DEPENDENCIES
    Requires Windows Forms.
    
    .VERSION
    1.0
    
    .AUTHOR
    Chris Thompson - christopher.thompson@leehealth.org
    
    .INSTRUCTIONS
    Call Get-FileName optionally specifying a custom window title and file filter. The function returns the full path to the selected file or $false if the operation is cancelled or an error occurs.
    
    .EXCEPTIONS
    Returns $false if the file dialog is cancelled, an error occurs, or the selected file does not exist.
    
    .ISSUES
    None known at this time.
    
    .TODO
    - Improve error handling for unexpected exceptions.
    - Consider adding support for multi-file selection.
    
    .NOTES
    This function is designed for scripts that require user input to select a file, offering a graphical interface that's intuitive for end-users.
    
    #>

    [CmdletBinding()]
    param ($title = "Select a CSV file", $filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*")

    # Ask the user to select a file
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = $filter
    $fileDialog.Title = $title

    try {
        # Launch form to promp user for file
        $fileDialogResult = $fileDialog.ShowDialog()
        $file = if ($fileDialogResult -eq [System.Windows.Forms.DialogResult]::OK) { $fileDialog.FileName } else { $false }

        Write-Debug "File chosen: $file"
    } catch {
        Write-Debug "Error in 'Get-Filename':`n$($Error[0])"
        $fileDialog.Dispose()
        return $false
    }
    if (Test-Path $file) {
        return $file
    } else {
        Write-Error "File not found"
        return $false
    }
}

# Function to load module and display errors if they are not found
function Load-ModuleGeneric {
    <#
    .SYNOPSIS
    Loads specified PowerShell modules, displaying errors for any that cannot be found.
    
    .DESCRIPTION
    Iterates through a list of modules and attempts to load each one. Displays an error message for any module that is not found or cannot be loaded.
    
    .PARAMETER ModuleList
    A hashtable of module names and optional paths to load. Each key represents a module name, and the value is another hashtable with 'Name', 'Path', and 'ErrorMessage'.
    
    .PARAMETER ErrorIdCounter
    An integer counter for tracking module load errors. Useful for scripts that load multiple modules and need to distinguish between different error messages.
    
    .EXAMPLE
    $moduleList = @{
        "Get-ComputerNames" = @{
            Name = "Get-ComputerNames"
            Path = $assetsFolder
        }
        "Active Directory" = @{
            Name = "ActiveDirectory"
            ErrorMessage = "This script requires the Active Directory module.`nYou need to install RSAT tools to proceed.`nRSAT Tools can be found at:`n\\lmh-sms\applications$\Microsoft\RSAT Tools"
        }
        "Progress Bar" = @{
            Name = "ProgressBar"
            Path = $assetsFolder
        }
    }
    Load-ModuleGeneric -ModuleList $moduleList
    
    .CHANGELOG
    1.0 Initial release.
    
    .DEPENDENCIES
    None.
    
    .VERSION
    1.0
    
    .AUTHOR
    Chris Thompson - christopher.thompson@leehealth.org
    
    .INSTRUCTIONS
    Define a hashtable with module details and pass it to this function to ensure all required modules are loaded before proceeding with script execution.
    
    .EXCEPTIONS
    Displays custom error messages for missing modules and increments an error ID counter for each module. Error M1 would mean the first module did not load.
    
    .ISSUES
    None.
    
    .TODO
    Implement automatic module installation if not found.
    
    .NOTES
    Customize the default error message and module path as needed.
    
    #>

    param (
        [Parameter(Mandatory)]
        [hashtable]$ModuleList,
        [int]$ErrorIdCounter = 1
    )

    foreach ($moduleKey in $ModuleList.Keys) {
        $module = $ModuleList[$moduleKey]
        $moduleName = $module.Name
        $modulePath = if ($module.Path) { $module.Path } else { $null }
        $errorMessage = if ($module.ErrorMessage) { $module.ErrorMessage } else { $null }
        if ($modulePath) { $moduleLoad = Join-Path $modulePath "$moduleName.psm1" }

        # Increment the global error ID counter
        $errorId = "M{0}" -f $ErrorIdCounter
        $ErrorIdCounter++
    
        # If no custom error message is provided, use a generic one
        if (-not $errorMessage) {
            $errorMessage = "The script requires the $moduleName module. Please ensure it is installed and available."
        }
    
        # Construct the module load path if $modulePath is provided
        if ($modulePath) {
            $moduleCondition = (Get-Module $moduleLoad -ListAvailable).ExportedCommands
        } else {
            $moduleCondition = (Get-Module $moduleName -ListAvailable).ExportedCommands
        }
    
        # Checks if the module is available
        if (!$moduleCondition) {
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Module Not Found", "OK", "Error")
            Write-Error "$errorMessage" -ErrorId $errorId -Category ResourceUnavailable
            $breakLoop = $true
            break
        }
    
        # Imports the module for use
        if ($modulePath) {
            Write-Debug "Module Path: $moduleLoad"
            Import-Module $moduleLoad -force #-WarningAction SilentlyContinue
        } else {
            Write-Debug "Module Name: $moduleName"
            Import-Module $moduleName -force #-WarningAction SilentlyContinue
        }
        Write-Debug "Loaded module: $moduleKey"
    }
    # Return the last ErrorIdCounter for future use (if called again or by other scripts launched by this one)
    if ($breakLoop) { return $false } else { return $ErrorIdCounter }
}