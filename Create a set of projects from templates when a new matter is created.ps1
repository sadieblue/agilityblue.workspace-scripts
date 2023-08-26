# ------------------------------------------------------------------------------
#
# When a user creates a new matter, this script will run through a list of project
# templates and create a new project for each template.
#
# In addition to automating project and task creation based on project templates,
# another benefit to using a script to create tasks is that the script can ignore
# required fields within the forms while creating the task. This allows you to
# create templates with required fields that are not required when the task is
# created from the template, but required when a user wants to update the task
# within the application.
#
# Event triggers to setup:
#   1) Object: Matter, Action: Create
#
# ------------------------------------------------------------------------------
# USER PARAMETERS

# Provide the list of project template IDs that the script will iterate through

$projectTemplateIds = @(
  2023080000003 # Matter Intake Project Template
  2023080000004 # Forensics Project Template
  2023080000005 # Processing Project Template
)

# ------------------------------------------------------------------------------
# SCRIPT FUNCTIONS

# Retrieves tasks that belong to the provided project id. 
# Returns a List<object> of task objects.
Function Get-TasksByProjectId([Int64]$projectId) {
  # Initialize a generic list
  $tasksToReturn = New-Object 'System.Collections.Generic.List[System.Object]'
  
  $taskTemplatesFilters = @(
    @{ Field = "ProjectId"; Value = $projectId }
  )
  
  # Use the Agility Blue `Get-Tasks` PowerShell command followed by filters to retrieve a
  # collection of tasks
  $taskTemplates = Get-Tasks -Filters $taskTemplatesFilters
  
  if ($null -eq $taskTemplates -or $taskTemplates.TotalCount -eq 0) {
    Write-Error "Template tasks for project $($projectId) could not be found"
    # Return an empty list if no tasks are found
    return , $tasksToReturn
  }
  
  # The tasks returned by the collection call only return simple task objects (tasks
  # without their forms). So we iterate through each task and retrieve the full task with
  # their forms.
  foreach ($simpleTask in $taskTemplates.Collection) {
    # Use the Agility Blue `Get-Task` PowerShell command followed by the task id and the
    # -IncludeForms switch parameter to retrieve the task that includes all of its forms
    $task = Get-Task $simpleTask.TaskId -IncludeForms
  
    # Add the task to the list
    $tasksToReturn.Add($task)
  }
  
  # Return the list of tasks. Note that we are using the comma operator to force PowerShell
  # to return a list, otherwise PowerShell will unroll single item lists as a single object.
  return , $tasksToReturn
}

# Creates a task from a template using the provided template task and project id.
# Returns a Task object.
Function Add-TaskFromTemplate($taskTemplate, [Int64]$projectId) {
  if ($null -eq $taskTemplate -or $null -eq $projectId) {
    Write-Error "taskTemplate and projectId are required"
    return
  }

  $taskTemplate.ProjectId = $projectId
  $taskTemplate.TaskId = 0

  # Loop through each form in the task template
  # The main purpose of looping through the forms is to zero out the IDs
  # that the template currently holds so the server doesn't try to associate
  # those IDs with the new task.
  for ($formIdx = 0; $formIdx -lt $taskTemplate.Forms.Count; $formIdx++) {
    $form = $taskTemplate.Forms[$formIdx]

    # Get the associated form the task template uses
    # Use the Agility Blue `Get-Form` PowerShell command followed by the form id to retrieve the form
    $templateForm = Get-Form $form.FormId

    # Zero out the IDs for the form
    $form.TaskFormId = 0
    $form.TaskId = 0

    # Loop through each section in the form
    for ($sectionIdx = 0; $sectionIdx -lt $form.Sections.Count; $sectionIdx++) {
      $section = $form.Sections[$sectionIdx]
      $templateSection = $templateForm.Sections[$sectionIdx]

      # Zero out the IDs for the section
      $section.TaskFormSectionId = 0
      $section.TaskFormId = 0
      $section.FormSectionId = $templateSection.SectionId

      # Loop through each field in the section
      for ($fieldIdx = 0; $fieldIdx -lt $section.Fields.Count; $fieldIdx++) {
        $field = $section.Fields[$fieldIdx]

        # Zero out the IDs for the field
        $field.TaskFormFieldId = 0
        $field.TaskFormSectionId = 0
      }
    }
  }

  # Create the task and return it back to the caller
  # Use the Agility Blue `Add-Task` PowerShell command followed by the task object to create the task
  $createdTask = Add-Task $taskTemplate

  return $createdTask
}

# ------------------------------------------------------------------------------
# SCRIPT BODY

# The $agilityBlueObject variable is a special object that Agility Blue populates
# if the script is triggered by an action. In the case of this script being
# triggered by a matter being created, the $agilityBlueObject variable will contain
# the matter object.
$matter = $agilityBlueObject

if ($null -eq $matter) {
  # If the event object is not available, it means we are executing the script 
  # manually. In this case, we'll get a matter directly for testing purposes.
  # Use the Agility Blue `Get-Matter` PowerShell command followed by the matter id to retrieve the matter
  $matter = Get-Matter 6089
    
  Write-Output "### Using test matter $($matter.MatterId) ###"
}

# Uncomment this line if you want to view the available properties/values for the matter
# $matter | ConvertTo-JSON -Depth 10

# Loop through each project template and create new projects based on each template
foreach ($projectTemplateId in $projectTemplateIds) {
  # use the Agility Blue `Get-Project` PowerShell command followed by the project id to retrieve the project template
  $projectTemplate = Get-Project $projectTemplateId

  # Get the template tasks for this template project using the custom function we defined earlier
  $taskTemplates = Get-TasksByProjectId -projectId $projectTemplateId

  # Uncomment this line if you want to view the task templates returned
  # $taskTemplates | ConvertTo-Json -Depth 10

  Write-Output "Applying project template $($projectTemplate.ProjectId) ($($projectTemplate.Description)) with $($taskTemplates.Count) task(s) to matter $($matter.MatterId) ($($matter.Name))..."

  # Create a new project object here using the project class we defined earlier
  $projectToCreate = @{
    MatterId                  = $matter.MatterId
    Description               = $projectTemplate.Description
    Priority                  = $projectTemplate.Priority
    EmailNotificationsEnabled = $projectTemplate.EmailNotificationsEnabled
  }

  # PowerShell will convert $null string properties to an empty string, so we'll
  # conditionally add the owned by id property if it is not null or empty, otherwise
  # The server throws an error.
  if (-not [string]::IsNullOrWhiteSpace($projectTemplate.OwnedById)) {
    $projectToCreate['OwnedById'] = $projectTemplate.OwnedById
  }

  # Use the Agility Blue `Add-Project` PowerShell command followed by the project object to create the project
  $createdProject = Add-Project $projectToCreate

  Write-Output "Created project $($createdProject.ProjectId) ($($createdProject.Description)) on matter $($createdProject.MatterId) ($($createdProject.MatterName))"

  # Loop through each task template and create a new task based on each template
  foreach ($taskTemplate in $taskTemplates) {
    # Wrap each call to create a task in a try-catch block so if one task fails it will move on to the next task.
    try {
      # Use the custom function we defined earlier to create a task from a template
      $createdTask = Add-TaskFromTemplate `
        -taskTemplate $taskTemplate `
        -projectId $createdProject.ProjectId
    
      Write-Output "Created template task $($createdTask.TaskId) ($($createdTask.Name)) on project $($createdTask.ProjectId) ($($createdTask.ProjectDescription))"
    }
    catch {
      Write-Error $_.Exception.Message
    }
  }
}