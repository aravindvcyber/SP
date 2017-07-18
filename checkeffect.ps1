param([switch]$help)
  
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Server")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Server.UserProfiles")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web")
function GetHelp() {
  
  
$HelpText = @"
  
DESCRIPTION:
This script will enumerate the permissions of the user in all webs under a site collection. This takes two input the user of the site collection and the username.The
  
username should be given in Domain\username format.
"@
$HelpText
  
}
  
function CheckEffectivePermissionsInAllWebs() {
  
write-host "This script will chcek the effective permissions of a user"
write-host "Please enter the url of the site collection"
$url = read-host
write-host "Please enter the username of the user"
$userName = read-host
$site = New-Object Microsoft.SharePoint.SPSite($url)
$serverContext = [Microsoft.Office.Server.ServerContext]::GetContext($site)
$userProfileManager = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($serverContext)
$userProfile = $userProfileManager.GetUserProfile($userName)
$userLogin = $userProfile[[Microsoft.Office.Server.UserProfiles.PropertyConstants]::AccountName].Value.ToString()
$webs = $site.AllWebs
foreach ($web in $webs)
{
$permissionInfo = $web.GetUserEffectivePermissionInfo($userLogin)
$roles = $permissionInfo.RoleAssignments
write-host "Now checking the permissions of the user "  $userLogin  " " "in the site " $web.Url
for ($i = 0; $i -lt $roles.Count; $i++)
{
$bRoles = $roles[$i].RoleDefinitionBindings
foreach ($roleDefinition in $bRoles)
{
 if ($roles[$i].Member.ToString().Contains('\'))
{
write-host "The User "  $userLogin  " has direct permissions "  $roleDefinition.Name
}
else
{
write-host "The User "  $userLogin  " has permissions "  $roleDefinition.Name  " given via "  $roles[$i].Member.ToString()
                                }
}
}
}
$site.Dispose()
}
  
if($help) { GetHelp; Continue }
else { CheckEffectivePermissionsInAllWebs }
