# set data to read only in file share. Only at level that is NOT INHERETED



$users = Import-csv ./.DS_Store
$domain = (Get-ADDomain).NetBIOSName


ForEach ($user in $users){
    $identity = "$domain\$($user.SamAccountName)"

    $Acl = Get-Acl -path $($user.Homefolder)
    $Ace = $Acl.Access | Where-Object {($_.IdentityReference -eq $identity) -and -not ($_.IsInherited)}
    $Acl.RemoveAccessRule($Ace)

    try {
    Set-Acl -Path $User.HomeFolder -AclObject $Acl
    Write-Host "Permissions removed for $identity on $($user.homefolder)"
    }
    catch {
        $Fault = $_
        Write-Host "Permissions cannot be removed for $identity on $($user.homefolder) : $($Fault.Exception.Message)"
    }

    $ace = $null 

    # new acl
    $rights = 'Read' #Other options: [enum]::GetValues('System.Security.AccessControl.FileSystemRights')
    $inheritance = 'ContainerInherit, ObjectInherit' #Other options: [enum]::GetValues('System.Security.AccessControl.Inheritance')
    $propagation = 'None' #Other options: [enum]::GetValues('System.Security.AccessControl.PropagationFlags')
    $type = 'Allow' #Other options: [enum]::GetValues('System.Securit y.AccessControl.AccessControlType')

    $ACE = New-Object System.Security.AccessControl.FileSystemAccessRule($identity,$rights,$inheritance,$propagation, $type)

    try {
        $Acl.AddAccessRule($ACE)
        Set-Acl -Path $User.HomeFolder -AclObject $Acl
        Write-Host "Permissions added for $identity on $($user.homefolder) to Read-Only"
    }
    catch {
        $Fault = $_
        Write-Host "Permissions cannot be added for $identity on $($user.homefolder) : $($Fault.Exception.Message)"
    }

}

