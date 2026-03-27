param (
  [string]$LogName = "System",
  [int]$EventID = 1001,
  [datetime]$StartTime = (Get-Date).AddDays(-1),
  [datetime]$EndTime = (Get-Date),
  [string]$LogFolder = "C:\Logs",
  [string]$OutputPath
)
 
# Define log folder and ensure it exists
if (-not (Test-Path -Path $LogFolder)) {
    try {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        Write-Output "Created folder: $LogFolder"
    } catch {
        throw "Failed to create folder $LogFolder. Error: $_"
    }
}
 
# Generate timestamp and computer name for unique file naming
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$computerName = $env:COMPUTERNAME
 
# Build output file path
if (-not $OutputPath) {
    $OutputPath = Join-Path $LogFolder "${computerName}_FilteredEventLogs_${timestamp}.csv"
}
 
try {
    Write-Output "Querying $LogName log for Event ID $EventID between $StartTime and $EndTime on $computerName..."
 
    $filter = @{ LogName = $LogName; Id = $EventID; StartTime = $StartTime; EndTime = $EndTime }
 
    $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop |
        Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, Message
 
    if (-not $events -or $events.Count -eq 0) {
        Write-Output "No events found for the given criteria."
    } else {
        $events | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
        Write-Output "Export complete. File saved to: $OutputPath"
    }
} catch [System.Exception] {
    if ($_.Exception.Message -like "*No events were found*") {
        Write-Output "No events found for the given criteria."
        exit 0
    }
    throw "An error occurred: $_"
}