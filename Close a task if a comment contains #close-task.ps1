# ------------------------------------------------------------------------------
#
# If a task comment is created that has the text #close-task, close the task.
# Useful for closing out tasks from email replies.
#
# Event triggers to setup:
#   1) Object: Task Comment, Action: Create
#
# ------------------------------------------------------------------------------

if ($null -eq $agilityBlueEvent) {
  # If the event object is not available, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $taskCommentId = 22178
} else {
  # The script was executed by an event, so we'll have access to the event object
  $taskCommentId = $agilityBlueEvent.Payload.Id
}

# Here, we are retrieving the comment and including the parent task so we can access 
# task properties.
$taskComment = Get-TaskComment $taskCommentId -IncludeTask

# We can uncomment this line to output the properties/values available to us
# $taskComment | ConvertTo-JSON

# First, we want to make sure that the task isn't already closed
if ($taskComment.Task.StatusDescription -in "Completed", "Canceled") {
  # The task is already closed out, so do nothing
  Write-Output "Task $($taskComment.TaskId) is $($taskComment.Task.StatusDescription). No action will be performed."
  Exit
}

# We're interested in the "Value" property here, specifically if it contains
# the text "#close-task"
if ($taskComment.Value -like "*#close-task*") {
  Write-Output "Comment $($taskComment.TaskCommentId) on task $($taskComment.TaskId) contains a #close-task command. Closing the task..."
  
  # Here, we issue a complete task command and set the Task property to the result
  $taskComment.Task = Set-CompleteTask $taskComment.TaskId
  
  # The Set-CompleteTask cmdlet will return a task object that we can use for output
  Write-Output "Task $($taskComment.TaskId) is now $($taskComment.Task.StatusDescription)"
} else {
  Write-Output "Comment $($taskComment.TaskCommentId) on task $($taskComment.TaskId) does not contain a #close-task command. No action will be performed."
}