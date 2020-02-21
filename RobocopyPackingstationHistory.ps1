# Local variables:
<# This section introduces local parameters that depend on 
the current machine (PC) and LAN. Edit them as needed.
#>
$destinationRootDir = "\\10.110.49.80\Packingstations\Pack12_CD";
$sourceRootDir = "C:\Emdep\LabelingStandalone\LabelingData\History";
$logFileName = "RobocopyPackingStationHistory.log";


function Write-AppEventLog {
 Param($entryType, $message)
 
    $eventSource = "RobocopyPackingstationHistory";
    $eventID = 65431;

    If ([System.Diagnostics.EventLog]::SourceExists($eventSource) -eq $False) {
       New-EventLog -LogName Application -Source $eventSource
    }

    # Writes an event to an event log. 
    Write-EventLog -LogName Application -EventID $eventID -EntryType $entryType `
        -source $eventSource -Message $message
}


function New-Directory {
 Param($itemPath)

    # If a destination directory doesn't exists then create its.
    if (-Not (Test-Path -Path $itemPath -PathType Container)) {
        New-Item -ItemType Directory -Path $itemPath;

        $msg = "Was created new directory (path): $itemPath";
        Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Information) $msg;
    }
}


function Copy-Data{
 Param($sourceDir, $destinationDir)
    
    $scriptLocation = Split-Path $MyInvocation.ScriptName;
    $logFile = ($scriptLocation, $logFileName) -join "\";
    
    # Copy all content including empty directory.
    Robocopy $sourceDir $destinationDir /s /np /r:1 /w:1 /log+:$logFile /copy:DAT

    $msg = "Copying file data executed successfully." + 
            "`nSource directory: $source" + 
            "`nDestination directory: $destination";

    Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Information) $msg;
}


function Clean-Up {
 Param($itemPath)

    # The number of deprecated objects which to be omitted.
    $number = 5;

    $itemsList = Get-ChildItem -path $itemPath | ?{$_.PSIsContainer} | Sort-Object -Property Name;
    
    if ($itemsList.Length -gt $number) {
        $deprecatedItemsList = $itemsList[0..($itemsList.Length - $number - 1)];
        $deprecatedItemsList | ForEach-Object {Remove-Item ($itemPath, $_ -join "\") -Recurse};
        
        $msg = "Deprecated items have cleaned successfully.";
        Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Information) $msg;
    }
}


function Main {
    # The number of objects of a list which to be selected.
    $number = 2;

    # The items list that will be copied.
    $itemsList = Get-ChildItem -Path $sourceRootDir | ?{$_.PSIsContainer} | `
        Sort-Object -Property CreationTime | Select-Object -Last $number;

    Start-Sleep -m 500;

    foreach ($item in $itemsList) {
        $destination = ($destinationRootDir, $item) -join "\";
        $source = ($sourceRootDir, $item) -join "\";
        New-Directory $destination;
        Start-Sleep -m 500;    
        Copy-Data $source $destination;
        Start-Sleep -m 500;
    }
      
    # Clean-up deprecated items.
    Clean-Up $destinationRootDir;
}


try {
    $msg = "Running the packing station history robocopy."
    Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Information) $msg;
    
    $dateToday = (Get-Date -Format yyyy_MM_dd).ToString();
    $newSource = ($sourceRootDir, $dateToday) -join "\";

    # Check if a source directory exists then do robocopy comand
    if (Test-Path -Path $newSource -PathType Container) {
        
        Main;
    }
    else {
        $msg = "Source directory ""$newSource"" doesn't exist. `nSo, nothing to copy.";
        Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Warning) $msg;
    }

    $msg = "Completed the packing station history robocopy."
    Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Information) $msg;
}
catch {
    $msg = "Some error occured: `n$_.Exception.Message";
    Write-AppEventLog ([System.Diagnostics.EventLogEntryType]::Error) $msg;
}
