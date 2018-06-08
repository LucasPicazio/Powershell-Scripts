$global:Summary = "This is the Summary"
$global:URI = "http://jiracorebr.csintra.net/"
$global:IssueParent = "Parent"
$global:ProjectKey = "ProjectKey"
$global:description = "This is the Issue Description"


$global:IssueCreationBody = 
@{
 fields = @{
 project = @{
 key = $projectKey
 }
 parent = @{
 key = $issueParent
 }
 summary = $summary
 description = $description 
 issuetype = @{
 id = '5' #this is a sub-task issuetype
 }
 assignee = @{
 name = $username
 }
 priority = @{
 id = '5' #no priority
 }
 }
 } | ConvertTo-Json -Depth 100
 #Try/catches are use so that if there is an error, it doesn't continue

 
 #create a session in JIRA in order to stay logged in and make changes
 try
 {
 Invoke-RestMethod -uri "$uri" -UseDefaultCredentials -ContentType "application/json" -Method post -Body $global:IssueCreationBody
 }
 catch
 {
 $fullerror = $_.Exception
 [void][System.Windows.Forms.MessageBox]::Show("Failed to log in with error:`r$fullerror", 'Error') # Casting the method to [void] suppresses the output. 
 $global:StopIssueCreation = $true
 }
 
 
 
 
 #this is a properly formed powershell JSON string
 #documentation on conversion can be found at 
 #<http://wahlnetwork.com/2016/02/18/using-powershell-arrays-and-hashtables-with-json-formatting/>
 <#
 $global:IssueCreationBody = 
@{
 fields = @{
 project = @{
 key = $projectKey
 }
 parent = @{
 key = $issueParent
 }
 summary = $summary
 description = $description 
 issuetype = @{
 id = '5' #this is a sub-task issuetype
 }
 assignee = @{
 name = $username
 }
 priority = @{
 id = '5' #no priority
 }
 }
 } | ConvertTo-Json -Depth 100
 
 
 if ($StopIssueCreation -ne $true) #if the login didn't fail
 {
 try
 {
 $IssueCreation = Invoke-WebRequest -uri "$uri/rest/api/latest/issue" `
 -ContentType "application/json" -Method post -Body $IssueCreationBody -WebSession $mysession | ConvertFrom-Json
 }
 
 catch
 {
 $fullerror = $_.Exception
 [void][System.Windows.Forms.MessageBox]::Show("Failed to create Issue in JIRA with error:`r$fullerror", 'Error') # Casting the method to [void] suppresses the output. 
 $global:StopIssueCreation = $true
 }
 }
 


<#Use this if you want to add a description Post-Creation
#and change the description above to ''

 if ($StopIssueCreation -ne $true)
 {
 
$global:descriptionBody = @{
 fields = @{ description = $description } 
} | ConvertTo-Json -Depth 100

 try
 {
 Invoke-WebRequest -uri "$uri/rest/api/latest/issue/$issueNumber" -ContentType "application/json" -Method put -websession $mysession -Body $descriptionBody | ConvertFrom-Json
$global:StopIssueCreation = $true
}
 catch
 {
 $fullerror = $_.Exception
 [void][System.Windows.Forms.MessageBox]::Show("Failed to add Description with error:`r$fullerror", 'Error') # Casting the method to [void] suppresses the output. 
 $global:StopIssueCreation = $true
 }
 }
 #>
 <#if ($StopIssueCreation -ne $true)
 {
 $global:repairCompleted = $true
 $global:StopIssueCreation = $false
 }
 if ($repairCompleted -ne $true)
 {} #If the repair didn't complete correctly, maybe tell them something #>