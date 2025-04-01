## You'll want the ImportExcel Module. You can import a CSV, but I prefer using this module.
## https://github.com/dfinke/ImportExcel
Import-Module ImportExcel


$InputFile = (Import-Excel -Path 'C:\temp\spreadsheet.xlsx').username ## Replace with applicable column
$FoundUsers = [System.Collections.ArrayList]@()
$ErrorUsers = [System.Collections.ArrayList]@()
Connect-MgGraph -Scopes User.Read.All


foreach ($user in $InputFile) {
    $mguser = Get-MgUser -Filter "UserPrincipalName eq '$user' or proxyAddresses/any(c:c eq 'smtp:$user')" -All
    if ($mguser) {
        [void]$FoundUsers.Add([PSCustomObject]@{
                Name           = $mguser.DisplayName
                AppUserName    = $user
                ActualUserName = $mguser.UserPrincipalName
                Match          = if ($user -eq $mguser.UserPrincipalName) { 'True' }else { 'False' }
            }
        )
    }
    else {
        [void]$ErrorUsers.Add([PSCustomObject]@{
                Name           = $user
                AppUserName    = $user
                ActualUserName = 'NOT FOUND!'
            }
        )
    }
}

$excelfile = "C:\temp\app_users_$(Get-Date -format yyyy-MM-dd).xlsx"
$FoundUsers | Export-Excel -WorksheetName 'FoundUsers' -Path $excelfile -BoldTopRow -AutoSize -AutoFilter
$ErrorUsers | Export-Excel -WorksheetName 'ErrorUsers' -Path $excelfile -BoldTopRow -AutoSize -AutoFilter

$excel = Open-ExcelPackage -Path $excelfile
$ws = $excel.Workbook.Worksheets["FoundUsers"]
$lastRow = $ws.Dimension.End.Row
#Matches
Add-ConditionalFormatting -Worksheet $excel."FoundUsers" -Range "C2:C$lastRow" -ConditionValue '=B2' -RuleType Equal  -BackgroundColor ([System.Drawing.Color]::PaleGreen) -Bold
Add-ConditionalFormatting -Worksheet $excel."FoundUsers" -Range "B2:B$lastRow" -ConditionValue '=C2' -RuleType Equal  -BackgroundColor ([System.Drawing.Color]::PaleGreen) -Bold
#Non Matches
Add-ConditionalFormatting -Worksheet $excel."FoundUsers" -Range "C2:C$lastRow" -ConditionValue '=B2' -RuleType NotEqual  -BackgroundColor ([System.Drawing.Color]::OrangeRed) -Bold
Add-ConditionalFormatting -Worksheet $excel."FoundUsers" -Range "B2:B$lastRow" -ConditionValue '=C2' -RuleType NotEqual  -BackgroundColor ([System.Drawing.Color]::OrangeRed) -Bold

Close-ExcelPackage -ExcelPackage $excel -Show
