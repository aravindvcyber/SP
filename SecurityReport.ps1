$ver = $host | select version
if($Ver.version.major -gt 1) {$Host.Runspace.ThreadOptions = "ReuseThread"}
if(!(Get-PSSnapin Microsoft.SharePoint.PowerShell -ea 0))
{
Add-PSSnapin Microsoft.SharePoint.PowerShell
}

##
#Set Script Variables
##

#Specify the Web Application to be targeted
$WebApplicationURL = "http://Contoso.com"

#Speciy the domain in which users exist, that you would want to report on
$UnwantedDomainPrefix = "BadDomain"

#Speicy a directory for which to log the results
$LoggingDirectory = "C:\PermissionReport\"

##
#Load Functions
##

#This function ensures that the logging directory exists before we start writing to any files
Function EnsureLoggingDirectory ($LoggingDirectory)
{
    #If the logging directory specified does not exist, create the directory
    if(!(Test-Path $LoggingDirectory))
    {
    #Create a file name dynamically based on the logging directory and the time of day that the script was executed
    Set-Variable -name Filename -Value ("$LoggingDirectory\SecurityReport_" +$StartTime +".txt") -Scope Script
    
    #Notify the user that the log file directory did not exist and that it will be created
    Write-Host "Path " $LoggingDirectory " does not exist. `r`nCreating Directory"
    
    #Create the logging directory
    New-Item -Path $LoggingDirectory -ItemType Directory
    }
}

#This function removes any trailing slashes from the logging directory, such that a dynamic filename can be created
Function TrimDirectory ($LoggingDirectory)
{
    #If the logging directory ends with a slash, proceed into the script block
    if($LoggingDirectory.EndsWith("\"))
    {
        #Remove the last character, which will be a slash, from the specified logging directory
        Set-Variable -Name LoggingDirectory -Value ($LoggingDirectory.Substring(0, ($LoggingDirectory.Length - 1))) -Scope Script
    }
}

##
#Start Script Execution
##

#Create a variable based on the current date and time
$StartTime = (Get-Date -UFormat "%Y-%m-%d_%I-%M-%S %p").tostring()

#Call the TrimDirectory function to remove trailing slashes from the logging directory
TrimDirectory $LoggingDirectory

#Call the EnsureLoggingDirectory function to ensure that the logging directory exists before we start writing to a log file
EnsureLoggingDirectory $LoggingDirectory

#Dynamically generate the log file name
$Filename = "$LoggingDirectory\SecurityReport_" +$StartTime +".txt"

#Create a brief header for the document, describing what information is in the report
"Security Report: $StartTime `r`n" | Out-File $Filename -Force
"Finding Users Beginning with "  + $UnwantedDomainPrefix + "\ and NTAuthority" | Out-File $Filename -Append

#Get all sites within each of the specified web application
$AllSites = Get-SPSite -WebApplication $WebApplicationURL -Limit All

#Perform some actions against each site retrieved
foreach($Site in $AllSites)
{
    #Return all webs in each site collection returned to the AllWebs variable
    $AllWebs = $Site.Allwebs
    
    #Write the target Site Collection to the log file 
    "`r`n`r`nSite: " +$Site.URL | Out-File $Filename -Append
    
    #Perform some actions against each web retrieved from the current target site collection
    foreach($Web in $AllWebs)
    {
        #Perfom some actions against a web if it has been assigned unique permissions
        if($Web.HasUniqueRoleAssignments)
        {
            #Determine if the target web is the root web of the site collection.  These sites are expected to have unique permissions assinged.
            if($Web.isrootweb)
            {
                #Log that the targeted web is the root web and will have unique permissions
                "`r`nweb '" + $Web.url + "' is Root Web and has unique permissions" | Out-File $Filename -Append
            }
            else
            {
                #For other webs which have unique permissions applied, simply log that he web uses unique permissions
                "`r`nWeb '" + $Web.url + "' is using unique permissions" | Out-File $Filename -Append
            }
            
            #Return a list of all role assignment entries at the web level
            $WebRoleAssignments = $Web.RoleAssignments
            
                #Perform some action against each role assignment returned
                foreach($WebRoleAssignment in $WebRoleAssignments)
                {
                    #Determine whether This is a an AD object or a SharePoint group.  Only domain users and groups have a userlogin property
                    if($WebRoleAssignment.member.userlogin)
                    {
                        #If the User's domain matches the target domain or NT Authority (for authenticatec users), perform some action
                        if($WebRoleAssignment.Member.userlogin.split("\")[0] -eq $UnwantedDomainPrefix -or $WebRoleAssignment.Member.userlogin.split("\")[0] -eq "NT Authority")
                        {
                            #Write an entry to the log file indicating that the targeted user has been granted a a specific right.
                            "User '" + $WebRoleAssignment.Member + "' has been assigned '" + ($WebRoleAssignment.RoleDefinitionBindings | select name).name + "'" | Out-File $Filename -Append
                        }
                    }
                    #For any entry which is determined to not be a user, it is a SharePoint group
                    else
                    {
                        #Return all users which have been added to the role assignment
                        $allWebUsers = $WebRoleassignment.member.users
                        
                        #Perform some action against all members returned.
                        foreach($WebUser in $AllWebUsers)
                        {
                            #If the User's domain matches the target domain or NT Authority (for authenticatec users), perform some action
                            if($WebUser.userlogin.split("\")[0] -eq $UnwantedDomainPrefix -or $WebUser.userlogin.split("\")[0] -eq "NT Authority")
                            {
                                #Write an entry to the log file indicating that the targeted user has been added to the specified SharePoint group.
                                "User '" + $WebUser.Userlogin + "' has been added to the '" + $WebRoleAssignment.member.name + "' group." | Out-File $Filename -Append
                            }
                        }
                    }
                }
            #Return a list of all lists in the target web
            $AllLists = $Web.lists    
            
            #Perform some action against all lists returned
            foreach($List in $AllLists)
            {
                #Perfom some actions against a list if it has been assigned unique permissions
                if($List.HasUniqueRoleAssignments)
                {
                    #Log to the log file thta the specified list is assigned u nique permissions
                    "`r`nList '" + $List.title + "' is using unique permissions `r`nURL: " + $WebApplicationURL +$List.DefaultViewURL | Out-File $Filename -Append
                    
                    #Return a list of all role assignment entries at the list level
                    $RoleAssignments = $List.RoleAssignments
                    
                    #Perform some action against all role assignments returned
                    foreach($RoleAssignment in $RoleAssignments)
                    {
                        #Determine whether This is a an AD object or a SharePoint group.  Only domain users and groups have a userlogin property
                        if($RoleAssignment.member.userlogin)
                        {
                            #If the User's domain matches the target domain or NT Authority (for authenticatec users), perform some action
                            if($RoleAssignment.Member.userlogin.split("\")[0] -eq $UnwantedDomainPrefix -or $RoleAssignment.Member.userlogin.split("\")[0] -eq "NT Authority")
                            {
                                #Write an entry to the log file indicating that the targeted user has been granted a a specific right.
                                "User '" + $RoleAssignment.Member + "' has been assigned '" + ($RoleAssignment.RoleDefinitionBindings | select name).name + "'" | Out-File $Filename -Append
                            }
                        }
                        #For any entry which is determined to not be a user, it is a SharePoint group
                        else
                        {
                            #Return all users which have been added to the role assignment
                            $allUsers = $Roleassignment.member.users
                            
                            #Perform some action against all members returned.
                            foreach($User in $AllUsers)
                            {
                                #If the User's domain matches the target domain or NT Authority (for authenticatec users), perform some action
                                if($user.userlogin.split("\")[0] -eq $UnwantedDomainPrefix -or $user.userlogin.split("\")[0] -eq "NT Authority")
                                {
                                    #Write an entry to the log file indicating that the targeted user has been added to the specified SharePoint group.
                                    "User '" + $User.Userlogin + "' has been added to the '" + $RoleAssignment.member.name + "' group" | Out-File $Filename -Append
                                }
                            }
                        }
                    }
                }
            }
        }
        else
        {
        $AllLists = $Web.lists
        "`r`n Web: " + $Web.url + " is inheriting permission from its Parent" | Out-File $Filename -Append    
            
            #Perform some action against all lists returned
            foreach($List in $AllLists)
            {
                #Perfom some actions against a list if it has been assigned unique permissions
                if($List.HasUniqueRoleAssignments)
                {
                    #Log to the log file thta the specified list is assigned u nique permissions
                    "`r`nList '" + $List.title + "' is using unique permissions `r`nURL: " + $WebApplicationURL +$List.DefaultViewURL | Out-File $Filename -Append
                    
                    #Return a list of all role assignment entries at the list level
                    $RoleAssignments = $List.RoleAssignments
                    
                    #Perform some action against all role assignments returned
                    foreach($RoleAssignment in $RoleAssignments)
                    {
                        #Determine whether This is a an AD object or a SharePoint group.  Only domain users and groups have a userlogin property
                        if($RoleAssignment.member.userlogin)
                        {
                            #If the User's domain matches the target domain or NT Authority (for authenticatec users), perform some action
                            if($RoleAssignment.Member.userlogin.split("\")[0] -eq $UnwantedDomainPrefix -or $RoleAssignment.Member.userlogin.split("\")[0] -eq "NT Authority")
                            {
                                #Write an entry to the log file indicating that the targeted user has been granted a a specific right.
                                "User '" + $RoleAssignment.Member + "' has been assigned '" + ($RoleAssignment.RoleDefinitionBindings | select name).name + "'" | Out-File $Filename -Append
                            }
                        }
                        #For any entry which is determined to not be a user, it is a SharePoint group
                        else
                        {
                            #Return all users which have been added to the role assignment
                            $allUsers = $Roleassignment.member.users
                            
                            #Perform some action against all members returned.
                            foreach($User in $AllUsers)
                            {
                                #If the User's domain matches the target domain or NT Authority (for authenticatec users), perform some action
                                if($user.userlogin.split("\")[0] -eq $UnwantedDomainPrefix -or $user.userlogin.split("\")[0] -eq "NT Authority")
                                {
                                    #Write an entry to the log file indicating that the targeted user has been added to the specified SharePoint group.
                                    "User '" + $User.Userlogin + "' has been added to the '" + $RoleAssignment.member.name + "' group" | Out-File $Filename -Append
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
