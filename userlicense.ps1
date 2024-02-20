# Import the Microsoft.Graph module
Write-Host "Importing Microsoft.Graph module..."
Import-Module Microsoft.Graph -WarningAction SilentlyContinue

# Suppress common warnings for the session
$WarningPreference = 'SilentlyContinue'

# Authenticate - this will prompt for login
Write-Host "Authenticating to Microsoft Graph..."
Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -WarningAction SilentlyContinue | Out-Null

# Retrieve users and their assigned licenses
Write-Host "Retrieving users and their assigned licenses..."
$users = Get-MgUser -All -Property id,userPrincipalName,assignedLicenses

# Initialize Excel application
Write-Host "Initializing Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)
$worksheet.Name = 'User Licenses'

# Create header row
$worksheet.Cells.Item(1, 1) = 'User Principal Name'
$columnIndex = 2

# Retrieve all licenses and prepare headers
Write-Host "Retrieving all available licenses..."
$allLicenses = Get-MgSubscribedSku
$licenseNames = @{}
$licenseCounts = @{}
foreach ($license in $allLicenses) {
    $licenseNames[$license.SkuId] = $license.SkuPartNumber
    $licenseCounts[$license.SkuPartNumber] = 0
    $worksheet.Cells.Item(1, $columnIndex) = $license.SkuPartNumber
    $columnIndex++
}

# Adjust row for user data start
$rowIndex = 3

foreach ($user in $users) {
    if ($user.UserPrincipalName -notlike "*#EXT#*" -and $user.AssignedLicenses.Count -gt 0) {
        Write-Host "Processing user: $($user.UserPrincipalName)"
        $worksheet.Cells.Item($rowIndex, 1) = $user.UserPrincipalName
        foreach ($licenseId in $user.AssignedLicenses.SkuId) {
            $licenseName = $licenseNames[$licenseId]
            $col = [array]::IndexOf($allLicenses.SkuId, $licenseId) + 2
            $cell = $worksheet.Cells.Item(2, $col) # Moving the total users count right below the license name
            $cell.NumberFormat = "@" # Ensure text format for centered alignment
            $cell.Value2 = ++$licenseCounts[$licenseName]
            $cell.HorizontalAlignment = -4108 # Center align in Excel
            $worksheet.Cells.Item($rowIndex, $col) = 'X'
            $worksheet.Cells.Item($rowIndex, $col).HorizontalAlignment = -4108 # Center align the 'X'
        }
        $rowIndex++
    }
}

# Formatting Excel document and saving
Write-Host "Formatting Excel document..."
$worksheet.Columns.AutoFit()
# Generate a unique GUID for the filename
$guid = [Guid]::NewGuid().ToString()
$filePath = "C:\temp\$guid.xlsx"
$workbook.SaveAs($filePath)
$excel.Quit()

# Clean up COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

# Disconnect from Microsoft Graph
Write-Host "Disconnecting from Microsoft Graph..."
Disconnect-MgGraph

Write-Host "Report generated successfully. File saved at: $filePath"

# Prompt user to open the file
$openFile = Read-Host "Do you want to open the report? (Y/N)"
if ($openFile -eq 'Y' -or $openFile -eq 'y') {
    Start-Process "EXCEL.EXE" $filePath
}
