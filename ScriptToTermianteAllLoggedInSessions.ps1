Import-Module MSOnline
$cred = Get-Credential
Connect-MsolService -Credential $cred
Connect-SPOService -Url https://yourdomainname-admin.sharepoint.com -Credential $cred

# To get all enabled users
#$users = Get-MsolUser -EnabledFilter EnabledOnly -All

# To get all the users of a Department
#$users = Get-MsolUser -Department 'Department Name' -All

# To get all the users
$users = Get-MsolUser -All;

foreach ( $user in $users){
Revoke-SPOUserSession -user $user.UserPrincipalName -Confirm:$false
}