# ------------------------------------------------------------------------------
#
# Update the priority of the containing project to "Rush" if the task 
#   due date is within 4 hours from the task created on date.
#
# ------------------------------------------------------------------------------
# [Post-Save Event Action]
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Task, Action: Create
#   2) Object: Task, Action: Update
#
# ------------------------------------------------------------------------------

#region Initialization ---------------------------------------------------------

$task = $agilityBlueObject

if ($null -eq $task) {
  # If the event object is not available, it means we are executing the script 
  # manually. In this case, we'll get a task directly for testing purposes.
  $task = Get-Task 29858
  
  Write-Output "### Using test task $($task.TaskId) ###"
}

# Uncomment this line to view the available properties/values
# $task | ConvertTo-JSON -Depth 10

#endregion

#region Functions --------------------------------------------------------------

Function Get-MinutesWithinDateRange([DateTime]$start, [DateTime]$end, [bool]$enforceBusinessHours) {
  [int]$minutes = 0

  for ($i = $start; $i -lt $end; $i = $i.AddMinutes(1)) {
    if ($enforceBusinessHours) {
      if (($i.TimeOfDay.Hours -ge 9) -and ($i.TimeOfDay.Hours -lt 18)) {
        $minutes++
      }
    }
    else {
      $minutes++
    }
  }

  return $minutes
}

# Tests that a task due date is within the specificed number of hours from now.
# If true, updates the project.
Function Test-DateAndUpdateProject($task, [int]$hoursFromDateCreated, [bool]$enforceBusinessHours) {
  # We convert all dates to UTC to ensure the date math works properly (dates are stored as UTC on the server, but do contain timezone information)
  $taskCreatedOnUtc = [DateTime]::Parse($task.CreatedOn).ToUniversalTime()
  $taskDateDueUtc = [DateTime]::Parse($task.DateDue).ToUniversalTime()

  $taskCreatedOnCst = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($taskCreatedOnUtc, "Central Standard Time")
  $taskDateDueCst = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($taskDateDueUtc, "Central Standard Time")

  $totalMinutes = Get-MinutesWithinDateRange `
    -start $taskCreatedOnCst `
    -end $taskDateDueCst `
    -enforceBusinessHours $enforceBusinessHours
  
  $totalHours = $totalMinutes / 60.0
  
  # Uncomment the next line to assist with debugging
  # Write-Output "Checking if the due date of $($taskDateDueCst) CST is less than $($hoursFromDateCreated) hours from the created on date of $($taskCreatedOnCst) CST (enforcing business hours: $($enforceBusinessHours) - $($totalMinutes) minutes || $($totalHours) hours)"

  if ($totalHours -le $hoursFromDateCreated) {
    # Conditions are satisfied - continue to mark this project as a rush
    
    Write-Output "Task $($task.TaskId) has a due date less than $($hoursFromDateCreated) hours from the created on date. Setting the priority to 'Rush' and saving the project"

    # Get the project
    $project = Get-Project $task.ProjectId

    # Set the priority
    $project.Priority = "Rush"

    # Update the project
    $updatedProject = Set-Project $project

    if ($null -eq $updatedProject) {
      Write-Error "There was an issue updating project $($project.ProjectId)"
    }

    # Check that the priority was indeed updated to "Rush"...
    if ($updatedProject.Priority -ne "Rush") {
      Write-Error "The server reported that the project was updated successfully, but the priority is reported to be set to $($project.Priority)"
    }

    Write-Output "The priority for project $($updatedProject.ProjectId) has been set to '$($updatedProject.Priority)'"
  }
  else {
    Write-Output "The due date for task $($task.TaskId) is not less than $($hoursFromDateCreated) hours from now"
  }
}

#endregion

#region Execution --------------------------------------------------------------

# We only care about normal projects...
if ($task.ProjectTypeName -ne "Normal") {
  Write-Output "Task $($task.TaskId) doesn't belong to a normal project type"
  Exit
}

# We only care about tasks that are open...
if ($task.StatusDescription -in "Completed", "Canceled") {
  Write-Output "Task $($task.TaskId) is closed"
  Exit
}

# If the task's project already has a "Rush" priority, there's no need to continue...
if ($task.Project.Priority -eq "Rush") {
  Write-Output "The priority for project $($task.ProjectId) is already set to 'Rush'"
  Exit
}

if ($null -eq $task.DateDue) {
  Write-Output "Task $($task.TaskId) does not have a due date"
  Exit
}

#Check if the task due date is less than 4 hours from the task created on date

Test-DateAndUpdateProject `
  -task $task `
  -hoursFromDateCreated 4 `
  -enforceBusinessHours $false

#endregion