# Trigger.ps1
# Update these two lines before use
$excelFilePath = "C:\Path\To\Your\File.xlsm"
$macroName = "CreateAndScanSubnets"

# Logging
$logDir = "C:\Path\To\Your\Logbook"
if (-not (Test-Path -Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logPath = "$logDir\Trigger_$(Get-Date -Format 'dd_MM_yyyy').txt"

function Write-TriggerLog($msg) {
    $timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    "[$timestamp] [PS-TRIGGER] $msg" | Out-File -FilePath $logPath -Append
}

Write-TriggerLog "Trigger script started."

$excel = $null
$wb = $null

try {
    Write-TriggerLog "Creating Excel COM object..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-TriggerLog "Opening file: $excelFilePath"
    $wb = $excel.Workbooks.Open($excelFilePath)

    Write-TriggerLog "Running macro: $macroName (this may take a while...)"
    $excel.Run($macroName)

    Write-TriggerLog "Macro completed. Saving and closing."
    $wb.Close($true)
}
catch {
    Write-TriggerLog "CRITICAL ERROR: $_"
}
finally {
    Write-TriggerLog "Cleanup started..."

    if ($wb -ne $null) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        Write-TriggerLog "Workbook object released."
    }

    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Write-TriggerLog "Excel closed and object released."
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-TriggerLog "Trigger script finished."
    Write-TriggerLog "--------------------------------------------------"
}
