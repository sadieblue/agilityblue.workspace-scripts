# ------------------------------------------------------------------------------
#
# This script will delete all normal/draft projects that don't have any tasks. 
# This can help cleanup workspaces where tasks are moved to other projects, but 
# the original project is left behind.
#
# ------------------------------------------------------------------------------
# [Manual Execution]
#
# ------------------------------------------------------------------------------

# Define a filter to find normal/draft projects that don't have any tasks
$noTasksFilter = "(ProjectTypeName eq 'Normal' or ProjectTypeName eq 'Draft') and NumberOfTasks eq 0"

# Retreive the list of projects that match the filter
$projectsToDelete = Get-Projects -Filter $noTasksFilter -Top 1000

if ($projectsToDelete.TotalCount -eq 0) {
  # No projects were found, exit the script
  Write-Output "No projects found with 0 tasks"
  Exit
}

Write-Output "Found $($projectsToDelete.TotalCount) projects with 0 tasks to delete"

$projectsRemoved = 0

# Iterate through each project and delete it
$projectsToDelete.Collection | ForEach-Object {
  $projectToDelete = $_

  try {
    Remove-Project -Id $projectToDelete.ProjectId
    $projectsRemoved++
  } catch {
    Write-Output "Failed to delete project $($projectToDelete.ProjectId): $_"
  }
}

Write-Output "Deleted $projectsRemoved projects"
