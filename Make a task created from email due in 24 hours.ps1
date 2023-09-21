# ------------------------------------------------------------------------------
#
# Make a task that was created from email due in 24 hours.
#
# ------------------------------------------------------------------------------
# [Post-Save Event Action]
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Task, Action: Create
#
# ------------------------------------------------------------------------------

if ($null -eq $agilityBlueEvent) {
  # If the event object is not available, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $taskId = 18764
}
else {
  # The script was executed by an event, so we'll have access to the event object
  $taskId = $agilityBlueEvent.Payload.Id
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
  ($null -eq $task.Metadata.Dictionary) -or 
  ($null -eq $task.Metadata.Dictionary.Source) -or 
  ($task.Metadata.Dictionary.Source -ne "Email")) {
  Write-Output "Task $($task.TaskId) was not created from email so a due date will not be automatically applied"
  Exit
}

# Set the due date to 24 hours from the created on date
$dueDate = $task.CreatedOn.AddHours(24)

# Set the due dates for the project and task
$task.Project.DateDue = $dueDate
$task.DateDue = $dueDate

# then save each one
$task.Project = Set-Project $task.Project
$task = Set-Task $task

Write-Output "Task $($task.TaskId) has been updated with a due date of $($task.DateDue) becuase it was created from email"
