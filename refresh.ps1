#VARIABLES
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$listOne= @(
    "e:\FangCloudV2\杭州初慕\初慕表格系统\库存分析\资料链接\采购分析\销量明细表\销量明细-2020.xlsx", 
    "e:\FangCloudV2\杭州初慕\初慕表格系统\库存分析\资料链接\采购分析\O2_销量分析.xlsx", 
    "e:\FangCloudV2\杭州初慕\初慕表格系统\库存分析\资料链接\采购分析\O2_销量图表.xlsm", 
    "e:\FangCloudV2\杭州初慕\初慕表格系统\库存分析\资料链接\采购分析\O3_追单日常.xlsx"
)
$listTwo= @(
    "e:\FangCloudV2\杭州初慕\初慕表格系统\库存分析\资料链接\O5_产品信息表.xlsm",
    "e:\FangCloudV2\杭州初慕\初慕表格系统\库存分析\资料链接\O3_下单明细表.xlsm"
)
$files = $listTwo
$files = $listOne

$Date = (Get-Date -Format 'yyyyMMdd-HHmm')
$errorFile = "C:\Temp\RefreshExcelError_" + $Date + ".txt" #Where you want an error file to be generated.
$isError = $false

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
        Write-Host $File" not found. `n" -foregroundcolor Red
		
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
		
        $excelObj = New-Object -ComObject Excel.Application
        $excelObj.Visible = $false

        #Open the workbook
        $workBook = $excelObj.Workbooks.Open($file)
        #Speed up calculation
        $excelObj.Calculation = $xlCalculationManual
		
        #Refresh all data twice in workbook.
        Write-Host "Starting refresh..." -nonewline
        $i = 1
        Do {
            $workBook.RefreshAll()
            $conn = $Workbook.Connections
            while ($conn | ForEach-Object { if ($_.OLEDBConnection.Refreshing) { $true } }) {
                Start-Sleep -Seconds 1
            }
            $i++
        } while ($i -le 1)
        Write-Host "done." 

        Write-Host "Starting Calculate..." -NoNewline
        $excelObj.Calculation = $xlCalculationAutomatic
        Write-Host "done."

        Write-Host "Saving file..." -nonewline
        $workBook.Save()
        Write-Host "done." 
		
        #Close workbook.
        $workBook.Close()
        $excelObj.Quit()
		
        #We must decrement the CLR reference count (to prevent the process from continuing to run in the background, which causes memory and lock problems).
        #https://technet.microsoft.com/en-us/library/ff730962.aspx
        #[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelObj)
        #Remove-Variable excelObj
        $excelObj = $null
        [GC]::Collect()
    }
    $end = Get-Date
    Write-Host -ForegroundColor Red ('Total Runtime: ' + ($end - $start).TotalSeconds)
    Write-Host "`n"
}
#Write-Host "`n"
#If an anticipated error found above, open the error file.
IF ($isError) {
    Invoke-Item $errorFile
}