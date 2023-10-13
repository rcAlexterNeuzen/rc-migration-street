
# Define script parameters
param(
    [Parameter(Mandatory)]
    [string]$Path,
    [Parameter(Mandatory)]
    [string]$outputPath
)

Import-Module ./modules/rc-required-modules.psm1

# Call the function to install required modules
Install-Requirements -modules "ImportExcel"

# Get a list of user shares/directories
$ADUsers = Get-ADUser -Filter *
$UserShares = Get-ChildItem -Path $Path -Directory

# Create an array for exporting user mapping data
$Export = @()
$Export | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value "" -force
$Export | Add-Member -MemberType NoteProperty -Name "userPrincipalName" -Value "" -force
$Export | Add-Member -MemberType NoteProperty -Name "Fullpath" -Value "" -force

# Create an array for storing user share sizes
$Size = @()
$Size | Add-Member -MemberType NoteProperty -Name "userPrincipalName" -Value "" -force
$Size | Add-Member -MemberType NoteProperty -Name "Fullpath" -Value "" -force
$Size | Add-Member -MemberType NoteProperty -Name "Size" -Value "" -force

# Process each user share
ForEach ($UserShare in $userShares){
    Write-Host "Processing $($Usershare.Name) " -NoNewline
    $AdInfo = $ADUsers | Where-Object {$_.SamAccountName -eq $UserShare.Name}
    if (!($AdInfo)){
        Write-Host "NOT FOUND" -ForegroundColor RED
    }
    Else {
        # Create an export item for user mapping
        $ExportItem = [PSCustomObject]@{
            DisplayName = $ADInfo.Name
            userPrincipalName = $ADInfo.userPrincipalName
            FullPath = $Usershare.FullName
        }
        
        # Create a size item for user share size
        $SizeItem = [PSCustomObject]@{
            userPrincipalName = $ADInfo.userPrincipalName
            FullPath = $Usershare.FullName
            Size = "{0:N2}" -f ((Get-ChildItem $userShare.FullName -recurse | Measure-Object Length -sum).sum / 1Mb)
        }

        Write-Host "DONE" -ForegroundColor Green
        
        # Add export and size items to their respective arrays
        $Export += $ExportItem
        $size += $SizeItem


    }
    
}

# Generate the filename for the output Excel file
$fileName = $outputPath + "\$(Get-Date -format "ddMMyyyy")-UserMapping-Homedrive.xlsx"

# Export user mapping data to Excel worksheet
$Export | Sort-Object -Property "DisplayName" | Export-Excel -WorksheetName "Homedrive Mapping" -Path $filename -FreezeTopRow -AutoFilter -AutoSize -BoldTopRow

# Export user share sizes to Excel worksheet
$size | Sort-Object -Property "userPrincipalName" | Export-Excel -WorksheetName "Size" -Path $filename -FreezeTopRow -AutoFilter -AutoSize -BoldTopRow
