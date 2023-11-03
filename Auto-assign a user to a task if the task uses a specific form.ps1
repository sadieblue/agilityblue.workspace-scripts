# ------------------------------------------------------------------------------
#
# Auto-assign a user to a task if the task uses a specific form
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Task, Action: On Create, Action State: After Save
#
# ------------------------------------------------------------------------------
# In order for this script to work properly, the $formName field must match the
# name of an existing form and the $assignToUserId needs to match a user id that
# exists within the workspace
#
# ------------------------------------------------------------------------------
# USER PARAMETERS

$formName = "Relativity Production"
$assignToUserId = "4895679e-ffe7-4097-9d92-a83cac5a68ae"

# ------------------------------------------------------------------------------

# The Get-InboundObjectInstance command will give you the object instance that was saved. 
# The type of object is dependent on the trigger context.
$task = Get-InboundObjectInstance

if ($null -eq $task) {
  # If the event object is not available, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $task = Get-Task 19820
}

# Uncomment this line to view the available properties/values
# $task | ConvertTo-JSON -Depth 10

if ($null -ne $task.AssignedToId) {
  # This indicates that there is already an assignee, so we'll do nothing
  Write-Output "Task $($task.TaskId) already has an assignee set"
  Exit
}

if ($task.FormNames -like "*$($formName)*") {
  # To update the assignee, we can't set the the AssignedToId directly on the task after the
  # task has already been created. Instead, we need to use the Set-AssignTask cmdlet

  # Save the assignee to the task. This will cmdlet will return the updated task
  $task = Set-AssignTask -TaskId $task.TaskId -UserId $assignToUserId

  Write-Output "Task $($task.TaskId) has been assigned to $($task.AssignedToFullName) (id: $($task.AssignedToId)) because the task is using the '$($formName)' form"
} else {
  Write-Output "Task $($task.TaskId) is not using the '$($formName)' form. Auto-assignment will not occur."
}