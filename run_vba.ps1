﻿param(
    [string]$files = $(throw "Parameter missing: -f File")
)
#VARIABLES
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105

$Date = (Get-Date -Format 'yyyyMMdd-HHmm')
$errorFile = "C:\Temp\RefreshExcelError_" + $Date + ".txt" #Where you want an error file to be generated.
$isError = $false

# function to close all com objects
function Release-Ref ($ref) {
    #([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

#Function to test filelock. Found on http://stackoverflow.com/questions/24992681/powershell-check-if-a-file-is-locked
function Test-FileLock {
    param (
        [parameter(Mandatory = $true)][string]$Path
    )
    $oFile = New-Object System.IO.FileInfo $Path
    if ((Test-Path -Path $Path) -eq $false) {
        return $false
    }

    try {
        $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($oStream) {
            $oStream.Close()
        }
        $false
    }
    catch {
        # file is locked by a process.
        return $true
    }
}

#Loop through all files, attempt to grab lock and refresh.
foreach ($file in $files) {
    #Ensure file exists
    IF (!(Test-Path $file)) {
        Write-Host $file" not found. `n" -foregroundcolor Red
		
        $errMsg = "FILE NOT FOUND: " + $file
        Add-Content $errorFile $errMsg
		
        $isError = $true;
        CONTINUE;
    }
    $start = Get-Date
    Write-Host $file -foregroundcolor Green
    Write-Host "Checking for lock in file..." -nonewline
	
    #Check if file is locked
    IF (Test-FileLock $file) {
        #File is locked.
		
        #Check if there is an error file yet.
        IF (!(Test-Path $errorFile)) { 
            #Error file doesn't exist, create one
            New-Item $errorFile -type file
        }
		
        #Add entry to log file
        $errMsg = "FILE LOCKED: " + $file
        Add-Content $errorFile $errMsg
        Write-Host "file locked." -foregroundcolor Magenta
        Write-Host "Error added to"+$errorFile -foregroundcolor Magenta		
        $isError = $true;
	
    }
    ELSE {	
        #File is NOT locked.
        Write-Host "file available."		
        $excelObj = New-Object -ComObject Excel.Application -ea Stop
        $excelObj.Visible = $true
        $excelObj.DisplayAlerts = $false
        $excelObj.ScreenUpdating = $false
        $workBook = $excelObj.Workbooks.Open($file)
        while (-not $excelObj.Ready) {
            Start-Sleep -Seconds 1
        }
        if ($file.Contains("库存分析.xlsm")) {
            $ws = $workBook.worksheets.item(1)
            Write-Host "Starting run macro..." -nonewline
            $excelObj.Run("Init")
            while ($excelObj.CalculationState -ne "xldone") {
                Start-Sleep -Seconds 5
            }
            Release-Ref($ws)
            Write-Host "done." 
        }
        else {    
            Write-Host "Starting refresh..." -nonewline
            $excelObj.Calculation = $xlCalculationManual #Speed up calculation    
            # foreach ($Sheet in $workBook.Worksheets) {
            #     foreach ($QTable in $Sheet.QueryTables) {
            #         $QTable.BackgroundQuery = $false
            #     }
            # }
            foreach ($cnn in $workBook.Connections) {
                if ($cnn.Type.ToString -eq "xlConnectionTypeODBC") {
                    $cnn.ODBCConnection.BackgroundQuery = $false              
                }
                else {
                    $cnn.OLEDBConnection.BackgroundQuery = $false
                }
            }
            $i = 1
            Do { 
                $workBook.RefreshAll()
                while ($workBook | ForEach-Object { if ($_.OLEDBCConnection.Refreshing) { $true } }) {
                    Start-Sleep -Seconds 1
                }
                $i++
            } while ($i -le 1) #Refresh all data twice in workbook.    
            foreach ($Sheet in $workBook.Worksheets) {
                foreach ($QTable in $Sheet.QueryTables) {
                    $QTable.BackgroundQuery = $true
                }
            }
            Write-Host "done." 

            Write-Host "Starting Calculate..." -NoNewline
            $excelObj.Calculation = $xlCalculationAutomatic
            while ($excelObj.CalculationState -ne "xldone") {
                Start-Sleep -Seconds 5
            }
            Write-Host "done."
        }
        Write-Host "Saving file..." -nonewline
        $workBook.Save()
        Write-Host "done." 
		
        #Close workbook.
        $workBook.Close()
        $excelObj.Quit()
        Get-Process -Name “*Excel*” | Stop-Process
        $excelObj = $null
    }
    ## close all object references
    
    Release-Ref($workBook)
    Remove-Variable workBook -ErrorAction SilentlyContinue
    Release-Ref($excelObj)
    Remove-Variable excelObj -ErrorAction SilentlyContinue
    $end = Get-Date
    Write-Host -ForegroundColor Red ('Total Runtime: ' + ($end - $start).TotalSeconds)
    Write-Host "`n"
}

#Write-Host "`n"
#If an anticipated error found above, open the error file.
IF ($isError) {
    Invoke-Item $errorFile
}