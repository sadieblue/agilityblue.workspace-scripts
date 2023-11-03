# ------------------------------------------------------------------------------
#
# Make a task that was created from email due in 24 hours.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Task, Action: On Create, Action State: After Save
#
# ------------------------------------------------------------------------------

$taskId = Get-InboundObjectId

if ($null -eq $taskId) {
  # If the Get-InboundObjectId returns null, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $taskId = 18764
}

# Note here that we want to get a full task that includes the project it belongs to. 
# The injected $agilityBlueObject will only provide a shallow version.
$task = Get-Task $taskId -IncludeProject

# Uncomment this line to view the available properties/values
# $task | ConvertTo-JSON -Depth 10

# We could check the $task.ProjectSourceTypeName property to see if the task was created
# from email, but this would also mean other tasks that gets created in this project would also
# automatically get a due date applied. Instead, we'll check for the metadata source parameter
# for email which is always set when a task is created from email at the task level.

if (($null -eq $task.Metadata) -or 
  ($null -eq $task.Metadata.Source) -or 
  ($task.Metadata.Source -ne "Email")) {
  Write-Output "Task $($task.TaskId) was not created from email so a due date will not be automatically applied"
  Exit
}

# Set the due date to 24 hours from the created on date
$dueDate = $task.CreatedOn.AddHours(24)

# Set the due dates for the project and task
$project = $task.Project;

$project.DateDue = $dueDate
$task.DateDue = $dueDate

# then save each one
$project = Set-Project $project
$task = Set-Task $task

Write-Output "Task $($task.TaskId) has been updated with a due date of $($task.DateDue) because it was created from email"
