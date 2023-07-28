# ------------------------------------------------------------------------------
#
# If a project is created that does not have the requester field set, this script
#   will use the matter default requester field to set it.
#
# Event triggers to setup:
#   1) Object: Project, Action: Create
#
# In order for this script to work properly, the Matter object must have a contact 
#   reference field named "Default Requester"
#
# ------------------------------------------------------------------------------

$project = $agilityBlueObject

if ($null -eq $project) {
  # If the event object is not available, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $project = Get-Project 2023070000001
}

# Uncomment this line to view the available properties/values
# $project | ConvertTo-JSON -Depth 10

# The "Requester" field is stored as "ContactId" on the project object
if ($null -ne $project.ContactId) {
  # This indicates that the user has already set a requester, so we'll do nothing
  Write-Output "Project $($project.ProjectId) already has a requester set"
  Exit
}

if ($null -eq $project.MatterId) {
  # This indicates that the project was created without a matter, so we'll do nothing
  Write-Output "Project $($project.ProjectId) does not belong to a matter"
  Exit
}

# Get the matter
$matter = Get-Matter $project.MatterId

# Get the "Default Requester" contact reference custom field
$defaultRequesterField = $matter.Fields | Where-Object { $_.Label -eq "Default Requester" }

# These are safety checks that determine if the matter does not have a "Default Requester" set
if (($null -eq $defaultRequesterField) -or
    ($null -eq $defaultRequesterField.Value) -or
    ($null -eq $defaultRequesterField.Value.ReferenceObject) -or
    ($null -eq $defaultRequesterField.Value.ReferenceObject.Values) -or
    ($defaultRequesterField.Value.ReferenceObject.Values.Count -eq 0)) {
  Write-Output "Matter $($matter.MatterId) does not have the 'Default Requester' field set"
  Exit
}

# Get the value of the "Default Requester" field. 
$defaultRequesterId = $defaultRequesterField.Value.ReferenceObject.Values[0].KeyAsInteger

# Set the project requester to the matter default requester
$project.ContactId = $defaultRequesterId

# Save the project
$project = Set-Project $project

# Log the results
Write-Output "The requester for project $($project.ProjectId) has been updated to $($project.ContactFullName) (id: $($project.ContactId))"
