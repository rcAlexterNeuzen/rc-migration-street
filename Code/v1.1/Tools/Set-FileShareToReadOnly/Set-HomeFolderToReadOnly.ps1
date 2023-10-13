# set data to read only in file share. Only at level that is NOT INHERETED
param(
    [Parameter(Mandatory)]
    [string]$CSV,
    [Parameter(Mandatory)]
    [string]$HomeFolder,
    [Parameter(Mandatory)]
    [string]$LogPath
)

if (!($isMacOs)) {
    $Switch = "\"
}
else {
    Write-Host "[ERROR] - This cannot be run on MacOs" -ForegroundColor Red
    break
}

## check if powershell elevated is started
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}
# version check
If ($PSVersionTable.PSVersion.Major -eq 5) {
    Write-Host "[INFORMATION] - Powershell $($PSVersionTable.PSVersion.Major) is present" -ForegroundColor Yellow
}
else {
    Write-Host "[INFORMATION] - Powershell $($PSVersionTable.PSVersion.Major) is present, setting throttle to 10" -ForegroundColor Yellow
    $Parallel = 10
}

# check if path HomeFolder exists
If (!(Test-Path $homefolder)) {
    Write-Host "[ERROR] - Homefolder base does not exists." -ForegroundColor red
    Break
}

Function Log-Message([String]$Message, [string]$file) {
    $message = "[$(Get-Date -format "HH:mm:ss")] " + $message
    Add-Content -Path $file $Message
}
function Remove-ACLPermission {
    param(
        [Parameter(Mandatory)]
        [string]$Identity,
        [Parameter(Mandatory)]
        [string]$Folder
    )

    $Acl = Get-Acl -path $Folder -ErrorAction Stop
    $Ace = $Acl.Access | Where-Object { ($_.IdentityReference -eq $identity) -and -not ($_.IsInherited) }
    forEach ($rule in $ace) {
        $Acl.RemoveAccessRule($rule) | out-Null
    }

    try {
        Set-Acl -Path $folder -AclObject $Acl -ErrorAction SilentlyContinue
        $output = "Permissions removed for $identity on $Folder"
    }
    catch {
        $Fault = $_
        $output = "Permissions cannot be removed for $identity on $folder : $($Fault.Exception.Message)"
    }

    return $output

}
function Add-ACLPermission {
    param(
        [Parameter(Mandatory)]
        [string]$Identity,
        [Parameter(Mandatory)]
        [string]$Folder,
        [Parameter(Mandatory)]
        [string]$Permission
    )
    try {
        $Acl = Get-Acl -path $folder -ErrorAction Stop
    }
    catch {
        $fault = $_
        $output = "ERROR getting ACL $Folder : $($Fault.Exception.Message)"
        break
    }

    $rights = $permission #Other options: [enum]::GetValues('System.Security.AccessControl.FileSystemRights')
    $inheritance = 'ContainerInherit, ObjectInherit' #Other options: [enum]::GetValues('System.Security.AccessControl.Inheritance')
    $propagation = 'None' #Other options: [enum]::GetValues('System.Security.AccessControl.PropagationFlags')
    $type = 'Allow' #Other options: [enum]::GetValues('System.Securit y.AccessControl.AccessControlType')

    $ACE = New-Object System.Security.AccessControl.FileSystemAccessRule($identity, $rights, $inheritance, $propagation, $type)

    try {
        $Acl.AddAccessRule($ACE)
        Set-Acl -Path $folder -AclObject $Acl
        $output = "Permissions added for $identity on $Folder to Read-Only"
    }
    catch {
        $Fault = $_
        $output = "Permissions cannot be added for $identity on $($user.homefolder) : $($Fault.Exception.Message)"
    }

    return $output
}

$file = $Logpath + $switch + "$(Get-Date -format "ddMMyyy-HHmm")-SetHomeFolderToReadOnly.txt"

Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "Start script for setting homefolders to Read-Only $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

Write-Host "----------------------------------------------"
Write-Host  "Start script for Installing Requirements $(Get-date -format "dd-MM-yyyy - HH:mm")"
Write-Host  "----------------------------------------------"
Write-Host  " "

# importing users
$Users = Import-CSV $CSV
$users | Add-Member -MemberType NoteProperty -Name "Status" -Value "" -force

If ($Parallel -eq 10) {
    # getting domain
    $Domain = $null
    Try {
        Import-Module ActiveDirectory -SkipEditionCheck -ErrorAction Stop
        $domain = (Get-ADDOmain).NetBIOSName
    }
    catch {
        if (!($Domain)) {
            $Domain = Read-Host "[QUESTION] - Please provide the NetBIOS name of the domain"
        }
    }

    $LogMessage = ${function:Log-Message}.tostring()
    $RemoveACLPermission = ${function:Remove-ACLPermission}.tostring()
    $AddACLPermission = ${function:Add-ACLPermission}.tostring()

    $users | Foreach-Object -ThrottleLimit $Parallel -Parallel {
        #Action that will run in Parallel. Reference the current object via $PSItem and bring in outside variables with $USING:varname
        ${function:Log-Message} = $using:LogMessage
        ${function:Remove-ACLPermission} = $using:RemoveACLPermission
        ${function:Add-ACLPermission} = $using:AddACLPermission
        ${Variable:file} = $using:file
        ${variable:domain} = $using:domain
        ${variable:switch} = $using:switch
        ${variable:homefolder} = $using:homefolder 


        $identity = "$domain\$($psitem.SamAccountName)"
        $path = $homefolder + $Switch + $($psitem.SamAccountName)
        if (!(Test-Path $path)) {
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -ForegroundColor Red -NoNewline
            Write-Host "] - $($Path) does not exist"
            Log-Message -file $file -Message "[ERROR] - $($Path) does not exist"
            $PSItem.Status = "ERROR - $($Path) does not exist"
        }
        else {
            try {
                $out = Remove-ACLPermission -Identity $identity -Folder $path
                Write-Host "[" -NoNewline
                Write-Host "DONE" -ForegroundColor Green -NoNewline
                Write-Host "] - $out"
                Log-Message -file $file -Message "[DONE] - $Out"
                $PSItem.Status = "DONE - $OUT"
            }
            catch {
                $Fault = $_
                Write-Host "[" -NoNewline
                Write-Host "ERROR" -ForegroundColor Red -NoNewline
                Write-Host "] - $out"
                Log-Message -file $file -Message "[ERROR] - $out"
                $PSItem.Status = "ERROR - $OUT"
            }

            try {
                $out = Add-ACLPermission -Identity $identity -Folder $path -Permission "Read"
                Write-Host "[" -NoNewline
                Write-Host "DONE" -ForegroundColor Green -NoNewline
                Write-Host "] - $out"
                Log-Message -file $file -Message "[DONE] - $out"
                $PSItem.Status = "DONE - $OUT"
            }
            catch {
                Write-Host "[" -NoNewline
                Write-Host "ERROR" -ForegroundColor Red -NoNewline
                Write-Host "] - $out"
                Log-Message -file $file -Message "[ERROR] - $out"
                $PSItem.Status = "ERROR - $OUT"
            }
        }
    }
}
else {
    # getting domain
    $Domain = $null
    Try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $domain = (Get-ADDOmain).NetBIOSName
    }
    catch {
        if (!($Domain)) {
            $Domain = Read-Host "[QUESTION] - Please provide the NetBIOS name of the domain"
        }
    }
    ForEach ($user in $users) {
        $identity = "$domain\$($user.SamAccountName)"
        $path = $homefolder + $Switch + $User.SamAccountName

        if (!(Test-Path $Path)) {
            Write-Host "[" -NoNewline
            Write-Host "ERROR" -ForegroundColor Red -NoNewline
            Write-Host "] - $Path does not exist"
            Log-Message -file $file -Message "[ERROR] - $Path does not exist"
            $User.Status = "ERROR - $Path does not exist"
        }
        else {
            try {
                $out = Remove-ACLPermission -Identity $identity -Folder $path
                Write-Host "[" -NoNewline
                Write-Host "DONE" -ForegroundColor Green -NoNewLine
                Write-Host "] - $out"
                Log-Message -file $file -Message "[DONE] - $Out"
                $User.Status = "DONE - $OUT"
            }
            catch {
                $Fault = $_
                Write-Host "[" -NoNewline
                Write-Host "ERROR" -ForegroundColor Red -NoNewLine
                Write-Host "] - $out"
                Log-Message -file $file -Message "[ERROR] - $out"
                $User.Status = "ERROR - $OUT"
            }
            try {
                $out = Add-ACLPermission -Identity $identity -Folder $path -Permission "Read"
                Write-Host "[" -NoNewline
                Write-Host "DONE" -ForegroundColor Green -NoNewLine
                Write-Host "] - $out"
                Log-Message -file $file -Message "[DONE] - $out"
                $User.Status = "DONE - $OUT"
            }
            catch {
                Write-Host "[" -NoNewline
                Write-Host "ERROR" -ForegroundColor Red -NoNewLine
                Write-Host "] - Permissions cannot be removed for $identity on $($Path) : $($Fault.Exception.Message)"
                Log-Message -file $file -Message "[ERROR] - Permissions cannot be removed for $identity on $($Path) : $($Fault.Exception.Message)"
                $User.Status = "ERROR - $OUT"
            }
        }
    }
}

$Users | Export-CSV $CSV 

Log-Message -file $file  " "
Log-Message -file $file "----------------------------------------------"
Log-Message -file $file  "End script for setting homefolders to Read-Only $(Get-date -format "dd-MM-yyyy - HH:mm")"
Log-Message -file $file  "----------------------------------------------"
Log-Message -file $file  " "

Write-Host  " " 
Write-Host "----------------------------------------------"
Write-Host  "End script for setting homefolders to Read-Only $(Get-date -format "dd-MM-yyyy - HH:mm")"
Write-Host  "----------------------------------------------"
Write-Host  " " 
 
