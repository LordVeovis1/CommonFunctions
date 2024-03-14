<#
.SYNOPSIS
This script provides a collection of common functions used across various scripts for system administration tasks.

.DESCRIPTION
CommonFunctions.ps1 contains reusable PowerShell functions including UI elements for data presentation, runtime measurement, asset verification, admin privilege checks, and module loading. These utilities facilitate script bootstrapping, ensuring prerequisites are met before script execution.

.PARAMETER bootstrap
Indicates whether the script should perform the bootstrap process, including loading necessary modules and verifying system requirements.

.EXAMPLE
# To use CommonFunctions.ps1 for bootstrapping:
$commonFunctions = "Path\To\CommonFunctions.ps1"
. $commonFunctions -bootstrap $true
This example demonstrates how to dot source CommonFunctions.ps1 with bootstrapping enabled.

.INPUTS
None. You cannot pipe objects to CommonFunctions.ps1.

.OUTPUTS
Output varies by function within the script. Generally, it includes console messages, form displays, and boolean values indicating the success or failure of specific checks.

.NOTES
Version:        1.0.0
Author:         Chris Thompson - christopher.thompson@leehealth.org
Creation Date:  03/07/24
Dependencies:   System.Windows.Forms assembly for UI elements.
Exceptions:     Various exceptions are handled within the script, particularly for file access, module loading, and privilege verification.
Issues:         No known issues at this time.

Instructions:   
- Dot source this script at the beginning of your PowerShell scripts to make its functions available.
- Adjust the $bootstrap parameter as necessary based on your script's requirements.

Updating:       
- Add new functions or modify existing ones as per evolving administrative tasks or script requirements.
- Ensure compatibility with existing scripts when updating.

Changelog:      
- 03/07/24: Initial creation. Introduced basic utilities for administrative scripting.

ToDo:
- Enhance user feedback and error handling for non-interactive environments.
- Enable optional bootstrap parameters, like admin = $true or $false to check for admin

Additional Notes:
This script is designed to be flexible and reusable across multiple administrative scripts. Modify and extend it as needed to fit your specific use cases.

#>

[CmdletBinding()]
param (
    $bootstrap = $true,
    $assetsFolder = '..\Script Assets'
)

Add-Type -AssemblyName System.Windows.Forms  # Load Windows Forms

#region This area is for showing a form based on an array of hashtables
# Define a custom sorter that implements the IComparer interface for sorting ListView items
class ListViewCustomSorter : System.Collections.IComparer {
    <#
    .SYNOPSIS
    Implements a custom sorter for ListView control items based on column text.
    
    .DESCRIPTION
    The ListViewCustomSorter class provides a sorting mechanism for items in a ListView control. It allows sorting based on the text of a specified column in either ascending or descending order. This class implements the IComparer interface necessary for custom sorting within ListView controls.
    
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
    
    .EXAMPLE
    $Data = @(
        @{ ComputerName = "PC1"; Status = "Online"; IP = "192.168.1.1" },
        @{ ComputerName = "PC2"; Status = "Offline"; IP = "192.168.1.2" }
    )
    Show-DataInForm -Data $Data

    This example shows how to display a list of computers with their status and IP addresses in a sortable form.

    .EXAMPLE
    $data = @(
        @{ Computer = "PC 1"; Status = "Online"; Latency = "93ms"; Devices = "4"; Monitors = "3 Monitor(s)"; IP = "192.168.1.118" }
        @{ Computer = "PC 2"; Status = "Unknown"; Latency = "85ms"; Devices = "3"; Monitors = "2 Monitor(s)"; IP = "192.168.1.94" }
        @{ Computer = "PC 3"; Status = "Unknown"; Latency = "37ms"; Devices = "2"; Monitors = "2 Monitor(s)"; IP = "192.168.1.227" }
        @{ Computer = "PC 4"; Status = "Online"; Latency = "9ms"; Devices = "3"; Monitors = "3 Monitor(s)"; IP = "192.168.1.238" }
        @{ Computer = "PC 5"; Status = "Offline"; Latency = "29ms"; Devices = "4"; Monitors = "3 Monitor(s)"; IP = "192.168.1.114" }
    )
    
    $Headers = @(Computer, Status, IP Address)
    
    Show-DataInForm -Data $data -ColumnHeaders $headers

    This example shows how to display a list of computers with their status and IP addresses in a sortable form.  This does not show the additional fields that are provided because of the headers included.

    #>
    
    param (
        [Parameter(Mandatory = $true)]
        $Data,  # Array of hashtables or PSCustomObjects, each representing a row of data.

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
        $computerKey = if ($ColumnHeaders -contains "ComputerName") { "ComputerName" } elseif ($ColumnHeaders -contains "Computer") { "Computer" } else { "ComputerName" }
        @($computerKey) + $ColumnHeaders.Where({ $_ -ne $computerKey })
    } else {
        # Check if the first item is a hashtable or a custom object and retrieve column names accordingly.
        if ($Data[0] -is [Hashtable]) {
            $keys = $Data[0].Keys
            $computerKey = if ($keys -contains "ComputerName") { "ComputerName" } elseif ($keys -contains "Computer") { "Computer" }
            @($computerKey) + $keys.Where({ $_ -ne $computerKey })

        # Check if the first item is a PSCustomObject
        } elseif ($Data[0].GetType().Name -eq 'PSCustomObject') {
            $keys = $Data[0].PSObject.Properties.Name
            $computerKey = if ($keys -contains "ComputerName") { "ComputerName" } elseif ($keys -contains "Computer") { "Computer" }
            @($computerKey) + $keys.Where({ $_ -ne $computerKey })
        } else {
            throw "Unsupported data type for Data parameter."
        }
    }

    foreach ($header in $columnsToAdd) {
        $listView.Columns.Add($header) | Out-Null
    }

    # Populate the ListView with data
    foreach ($item in $Data) {
        $values = $columnsToAdd | ForEach-Object {
            if ($item -is [Hashtable]) {
                $item[$_]
            } elseif ($item.GetType().Name -eq 'PSCustomObject') {
                $item.PSObject.Properties[$_].Value
            } else {
                throw "Unsupported data type in Data array."
            }
        }

        $listViewItem = New-Object System.Windows.Forms.ListViewItem($values[0].ToString())
        $values[1..($values.Count - 1)] | ForEach-Object {
            $subItemText = if ($_ -eq $null) { "" } else { $_.ToString() }
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
    Measures and displays the runtime of a PowerShell script.
    
    .DESCRIPTION
    The `Measure-ScriptRuntime` function calculates and displays the elapsed time for a script's execution. It can mark the script's start time and calculate the elapsed time upon completion. If a start time is provided, it calculates and displays the elapsed time since that point. If no start time is provided, the current time is returned, useful for marking the execution start.
    
    .PARAMETER startTime
    (Optional) Specifies the start time from which the script's runtime is calculated. If omitted, the function returns the current time, which can be used as a start marker for execution.
    
    .EXAMPLE
    # Mark the start of the script:
    $startTime = Measure-ScriptRuntime
    
    # Script operations here
    
    # Calculate and display script runtime at the end:
    Measure-ScriptRuntime -startTime $startTime
    
    Demonstrates marking the start of a script's execution and measuring its total runtime at the end.
    
    .EXAMPLE
    # Mark the start of the first operation:
    $firstOperationStart = Measure-ScriptRuntime
    
    # First operation code here
    
    # Calculate and display the first operation's runtime:
    Measure-ScriptRuntime -startTime $firstOperationStart
    
    # Mark the start of the second operation:
    $secondOperationStart = Measure-ScriptRuntime
    
    # Second operation code here
    
    # Calculate and display the second operation's runtime:
    Measure-ScriptRuntime -startTime $secondOperationStart
    
    This example shows how to time separate sections of a script independently, offering detailed insights into the execution time of each part.
    
    .NOTES
    
    This function is invaluable for performance monitoring and script optimization, providing a straightforward approach to timing script execution without complex mechanisms or external tools.
    
    #>

    param (
        [DateTime]$startTime
    )

    if ($startTime) {
        # Calculate and display the elapsed time
        $endTime = Get-Date                                                              # Capture the current time as the end time
        $elapsedTime = $endTime - $startTime                                             # Calculate elapsed time since start
    
        $elapsedTimeSpan = [TimeSpan]::FromMilliseconds($elapsedTime.TotalMilliseconds)  # Convert elapsed time to a TimeSpan object for easy formatting
        $formattedElapsedTime = $elapsedTimeSpan.ToString("hh\:mm\:ss")                  # Format as hh:mm:ss
    
        Write-Host "Script finished at: $endTime"                                        # Display end time
        Write-Host "Total script runtime: $formattedElapsedTime"                         # Display formatted elapsed time
    } else {
        $startTime = Get-Date                                                            # No start time provided, use current time as start
        Write-Host "Script started at: $startTime"                                       # Inform the user script start time
        return $startTime                                                                # Return the start time for potential future use
    }
}

# Function to verify Assets folder location and set directory location for script execution
function Verify-AssetsFolder {
    <#
    .SYNOPSIS
    Verifies the presence of an Assets folder and a specific CSV file.
    
    .DESCRIPTION
    The function checks for the existence of a specified Assets folder and a CSV file named "Building and Department Codes.csv" within it. If the folder is not found, the user is prompted to select a folder through a graphical interface. The function ensures that critical assets are available before proceeding with the script execution.
    
    .PARAMETER assetsFolder
    Specifies the path to the Assets folder. If not provided, a default path is used. The function checks this folder for the presence of the required CSV file.
    
    .EXAMPLE
    # To verify the default Assets folder:
    Verify-AssetsFolder
    
    .EXAMPLE
    # To specify a custom Assets folder path:
    Verify-AssetsFolder -assetsFolder "C:\Custom\Path\To\Assets"
    
    .OUTPUTS
    Boolean
    Returns $true if the Assets folder and the CSV file are found. Returns $false and terminates script execution otherwise.
    
    .NOTES
    Dependencies: System.Windows.Forms assembly for displaying the folder browser dialog.
    
    #>

    param(
        [string]$assetsFolder = "\\lmh-sms\utilities$\Script Assets"  # Define the default Assets folder path
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
    
    # Ensure the Building and Department Codes CSV file exists within the Assets folder
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
            Import-Module $moduleLoad -WarningAction SilentlyContinue
        } else {
            Write-Debug "Module Name: $moduleName"
            Import-Module $moduleName -WarningAction SilentlyContinue
        }
        Write-Debug "Loaded module: $moduleKey"
    }
    # Return the last ErrorIdCounter for future use (if called again or by other scripts launched by this one)
    if ($breakLoop) { return $false } else { return $ErrorIdCounter }
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

# Function for Save File Dialog - default .csv
function Show-SaveFileDialog {
<# How to use the function
.SYNOPSIS
    Presents a Save File dialog to the user and returns the selected file path.

.DESCRIPTION
    This function shows a graphical user interface dialog for saving a file. 
    It allows the user to specify the location and name of the file to save. 
    The function parameters enable customization of the dialog, 
    including setting the file type filter, dialog title, and handling user cancellation.

    If the user selects a file and confirms, the function returns the file path.
    If the user cancels the dialog, a message box is shown, an error is logged, 
    and the script exits.

.PARAMETER filter
    Specifies the file type filter for the dialog. Default is "CSV Files (*.csv)|*.csv".

.PARAMETER title
    Sets the title of the save file dialog. Default is "Save data results to CSV".

.PARAMETER message
    The message displayed in a message box if no file is selected. 
    Also used for the error message. Default is "No file selected. Data cannot be saved.`nExiting script".

.PARAMETER errorId
    The identifier for the error message if the user cancels the dialog. 
    Default is "FileSaveCancelled".

.EXAMPLE
    $filePath = ShowSaveFileDialog -filter "Text Files (*.txt)|*.txt" -title "Save Text File"

    This example shows the save file dialog with a filter for text files and a custom title.
#>
    param (
        [string]$filter  = "CSV Files (*.csv)|*.csv",
        [string]$title   = "Save data results to CSV",
        [string]$message = "No file selected. Data cannot be saved.`nExiting script",
        [string]$errorId = "FileSaveCancelled"
    )

    # Create and configure the SaveFileDialog
    $saveFileDialog        = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = $filter
    $saveFileDialog.Title  = $title

    # Show the dialog and process the user's response
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $outputFilePath = $saveFileDialog.FileName
    } else {
        [System.Windows.Forms.MessageBox]::Show($message, "No File Selected", "OK", "Error")
        Write-Error $message -ErrorId $errorId -Category ObjectNotFound
        return $false
    }

    # Clean up resources
    $saveFileDialog.Dispose()

    # Return the selected file path
    return $outputFilePath
}

# Function to export user data to CSV
function Export-UserDataToCSV {
    param (
        [Parameter(Mandatory = $true)]
        $OrderedUserDataArray,

        [Parameter(Mandatory = $true)]
        [string]$outputFilePath
    )

    try {
        # Export the collected data to the selected file with headers
        $OrderedUserDataArray | Export-Csv -Path $outputFilePath -NoTypeInformation

        # Inform the user that the data has been successfully saved
        [System.Windows.Forms.MessageBox]::Show("Data successfully saved to $outputFilePath", "Export Successful", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        Write-Host "Data saved to $outputFilePath"
    } catch {
        # Inform the user that there was an error saving the data
        [System.Windows.Forms.MessageBox]::Show("An error occurred while saving the data to $outputFilePath. Please check the file path and try again.", "Export Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Write-Host "Failed to save data to $outputFilePath"
    }
}

# Function to prompt how to return the data to the user
function Get-DataFormat {
    # Create the form
    $form                       = New-Object System.Windows.Forms.Form
    $form.Text                  = 'Select Action'
    $form.Size                  = New-Object System.Drawing.Size(300, 200)

    # Option 1: Save to File
    $radioSaveToFile            = New-Object System.Windows.Forms.RadioButton
    $radioSaveToFile.Location   = New-Object System.Drawing.Point(10, 10)
    $radioSaveToFile.Size       = New-Object System.Drawing.Size(200, 20)
    $radioSaveToFile.Text       = 'Save to File'

    # Option 2: Display Popup
    $radioDisplayPopup          = New-Object System.Windows.Forms.RadioButton
    $radioDisplayPopup.Location = New-Object System.Drawing.Point(10, 40)
    $radioDisplayPopup.Size     = New-Object System.Drawing.Size(200, 20)
    $radioDisplayPopup.Text     = 'Display Popup'

    # Option 3: Both
    $radioBoth                  = New-Object System.Windows.Forms.RadioButton
    $radioBoth.Location         = New-Object System.Drawing.Point(10, 70)
    $radioBoth.Size             = New-Object System.Drawing.Size(200, 20)
    $radioBoth.Text             = 'Both'
    $radioBoth.Checked          = $true

    # OK Button
    $okButton                   = New-Object System.Windows.Forms.Button
    $okButton.Location          = New-Object System.Drawing.Point(100, 100)
    $okButton.Size              = New-Object System.Drawing.Size(75, 23)
    $okButton.Text              = 'OK'
    $okButton.DialogResult      = [System.Windows.Forms.DialogResult]::OK

    
    # Add controls to form
    $form.Controls.Add($radioDisplayPopup)  # Add a radio option for Popup
    $form.Controls.Add($radioSaveToFile)    # Add the radio option for save-to-file
    $form.Controls.Add($radioBoth)          # Add a radio option for both
    $form.Controls.Add($okButton)           # Add the OK button
    
    # Set the OK button as the default for the enter button
    $form.AcceptButton = $okButton

    # Show the form
    $form.Topmost = $true
    $form.ShowDialog()

    # Check which option was selected and return the result
    if ($radioSaveToFile.Checked) {
        $form.Dispose()
        return 'SaveToFile'
    } elseif ($radioDisplayPopup.Checked) {
        $form.Dispose()
        return 'DisplayPopup'
    } elseif ($radioBoth.Checked) {
        $form.Dispose()
        return 'Both'
    }
}

# Execute a boostrap proceedure, if specified (default is $true, use '-Bootstrap $false' to prevent bootstrap execution)
if ($bootstrap) {
    try {
        #region This section requires no user interaction
        # Variables
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
    } catch {
        throw "Bootstrap failed.`n$Error[0]"
    }
}
