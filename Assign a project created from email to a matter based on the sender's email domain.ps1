# ------------------------------------------------------------------------------
#
# When an email is sent into Agility Blue, this script will check the sender's 
# email domain and search for a matter that contains that domain in a custom text 
# field. If a match is found, the task's project will be updated to use the 
# matching matter.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Task, Action: On Create, Action State: After Save
#
# ------------------------------------------------------------------------------
# In order for this script to work properly, the Matter object must have a basic
# text field that can hold email domains. The field id of this field must be 
# provided in the $emailDomainsFieldId variable below.
#
# Note: The name of the basic text field doesn't matter in the context of this
# script. The script will use the field id to retrieve the field's value. We
# suggest calling it "Email Domains" or something similar.
#
# Note: The field can have more than one email domain because the script uses a 
# contains filter to find a match. The domains can be separated in the field
# however you like.
#
# ------------------------------------------------------------------------------
# USER-DEFINED VARIABLES
#
# The id of the custom text field on the matter object that contains email domains
$emailDomainsFieldId = 70
#
# ------------------------------------------------------------------------------

# The `Get-InboundObjectInstance` command will provide scripts with the incoming
# object instance based on the trigger object. In this case, the trigger object
# is a task.
$task = Get-InboundObjectInstance

if ($null -eq $task) {
  # If the inbound object instance is null, it means the script is being executed
  # manually. In this case, we'll retrieve a specific task for testing purposes.
  $task = Get-Task -Id 935
  
  Write-Output "*** Using test task $($task.TaskId) ***"
}

# We check on the task's metadata to see if it originated from an email.
# The metadata property is not present in every task instance, but this information is
# always added by the email system for every email that comes into Agility Blue.
if (($null -eq $task.Metadata) -or
    ($null -eq $task.Metadata.Source) -or
    ($null -eq $task.Metadata.From) -or
    ($task.Metadata.Source -ne "Email")) {
  # Because the task did not originate from an email, we'll skip processing the rest 
  # of the script and exit.
  Write-Output "Task $($task.TaskId) did not originate from an email. Skipping further processing."
    
  exit
}

Write-Output "Task $($task.TaskId) originated from an email"

# This portion of the script extracts the domain from the sender's email address.
# We'll use the domain to search for matching matters.
$fromEmailParts = $task.Metadata.From.Split("@")

if ($fromEmailParts.Length -ne 2) {
  Write-Output "Expected the 'from' email domain to contain 2 parts, but it contained $($fromEmailParts.Length)"
    
  exit
}

$fromDomain = $fromEmailParts[1]

if ([String]::IsNullOrWhiteSpace($fromDomain)) {
  Write-Output "Expected the 'from' email domain to have a value"
    
  exit
}

Write-Output "Searching matters for a matching email domain containing '$($fromDomain)' for task $($task.TaskId)..."

# We can apply an odata filter on matters that contain our $fromDomain to see if we have a match
# Refer to the API documentation for information about odata and filtering.
$matchingMatters = Get-Matters -Filter "contains(CF_$($emailDomainsFieldId),'$($fromDomain)')"

if ($matchingMatters.TotalCount -eq 0) {
  Write-Output "No matters matched an email domain containing '$($fromDomain)' for task $($task.TaskId). No changes made."
    
  exit
}

# A match was made, grab the first matter from the Collection
$matchingMatter = $matchingMatters.Collection | Select-Object -First 1

Write-Output "The email domain '$($fromDomain)' matched with matter $($matchingMatter.MatterId) for task $($task.TaskId)"

# Retrieve the task's parent project using the `Get-Project` command
$parentProject = Get-Project -Id $task.ProjectId

# Update the parent project's matter to use the matching matter
$parentProject.MatterId = $matchingMatter.MatterId

# Save the updated project using the `Set-Project` command
$updatedProject = Set-Project $parentProject

# Write information to the output stream that will be visible in the execution
# results logs
Write-Output "Updated project $($updatedProject.ProjectId) to use matter $($updatedProject.MatterId)"
