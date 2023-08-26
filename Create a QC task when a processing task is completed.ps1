# ------------------------------------------------------------------------------
#
# When a user completes a task, this script will create a new task on the same
# project that uses a QC form.
#
# In addition to automating project and task creation based on forms,
# another benefit to using a script to create tasks is that the script can ignore
# required fields within the forms while creating the task. This allows you to
# create forms with required fields that are not required when the task is
# created from the form, but required when a user wants to update the task
# within the application.
#
# Event triggers to setup:
#   1) Object: Task, Action: Update
#
# ------------------------------------------------------------------------------
# USER PARAMETERS

# The task form id is checked when a task is completed. If the task contains a form
# with this id, then a new task will be created using the QC form.
$taskFormId = 16374 # Processing Form

# The QC form id is used to create a new task when a task is completed, provided the
# task that was completed contains the taskFormId from above.
$qcFormId = 6122 # Processing QC Form

# ------------------------------------------------------------------------------
# SCRIPT BODY

# The $agilityBlueEvent variable is a special object that Agility Blue populates
# if the script is triggered by an event. In the case of this script being
# triggered by a task being updated, the $agilityBlueEvent variable will be
# available.
if ($null -eq $agilityBlueEvent) {
  # If the event object is not available, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $taskId = 29952
}
else {
  # The script was executed by an event, so we'll have access to the event object
  # Here we get the task id from the event objects, payload id property.
  $taskId = $agilityBlueEvent.Payload.Id
}

# Here, we are retrieving the task and including the forms so we can access 
# form properties.
$task = Get-Task $taskId -IncludeForms

# We can uncomment this line to output the properties/values available to us
# $task | ConvertTo-JSON

# We want to make sure that the task is closed
if ($task.StatusDescription -ne "Completed") {
  # The task is already closed out, so do nothing
  Write-Output "Task $($task.TaskId) is not completed. No action will be performed."
  Exit
}

# Check if the task contains the task form id we are looking for
if (-not ($task.Forms | Where-Object { $_.FormId -eq $taskFormId })) {
  # The task doesn't contain the form we're looking for, so do nothing
  Write-Output "Task $($task.TaskId) doesn't contain form $($taskFormId). No action will be performed."
  Exit
}

# Before we create a QC task, let's make sure that one doesn't already 
# exists in the project.

$qcTaskAlreadyExists = $false

$tasksInProjectFilters = @(
  @{ Field = "ProjectId"; Operator = "="; Value = $task.ProjectId }
)

$tasksInProject = Get-Tasks -Filters $tasksInProjectFilters -Top 1000

$tasksInProject.Collection | ForEach-Object {
  $taskInProject = Get-Task $_.TaskId -IncludeForms

  if ($taskInProject.Forms | Where-Object { $_.FormId -eq $qcFormId }) {
    $qcTaskAlreadyExists = $true
    return
  }
}

if ($qcTaskAlreadyExists) {
  # A QC task already exists in the project, so do nothing
  Write-Output "A QC task already exists in project $($task.ProjectId). No action will be performed."
  Exit
}

# Get the form that we want to use for the QC task
# Use the Agility Blue `Get-Form` PowerShell command followed by the 
# form id to retrieve the form
$qcForm = Get-Form $qcFormId

# Create a new task wiht the basic task information
$qcTask = @{
  Name      = $qcForm.Name
  ProjectId = $task.ProjectId
  DateDue   = $task.DateDue
  Forms     = @()
}

# To use a form for a task, we need to create a task form object that's
# based on the form we want to use. We'll loop through each section and
# field in the form and create a task form object that we can use to
# create the task.

# Forms contain sections, and sections contain fields.

$qcTaskForm = @{
  FormId   = $qcForm.FormId
  Sections = @()
}

$qcForm.Sections | ForEach-Object {
  $qcFormSection = $_

  $qcTaskFormSection = @{
    FormSectionId = $qcFormSection.SectionId
    Fields        = @()
  }

  $qcTaskForm.Sections += $qcTaskFormSection

  $qcFormSection.Fields | ForEach-Object {
    $qcFormSectionField = $_

    $qcTaskFormSectionField = @{
      FieldId = $qcFormSectionField.FieldId
    }

    # Check if there is a default value on the field. If there is, apply the default
    # To the appropriate ValueAs* field
    if (-not [string]::IsNullOrWhiteSpace($qcFormSectionField.DefaultValue)) {
      switch ($qcFormSectionField.DataTypeName) {
        "Whole Number" {
          # Convert the DefaultValue to an integer
          $qcTaskFormSectionField['ValueAsNumber'] = `
            [int]$qcFormSectionField.DefaultValue
        }
        "Decimal Number" {
          # Convert the DefaultValue to a decimal
          $qcTaskFormSectionField['ValueAsDecimal'] = `
            [decimal]$qcFormSectionField.DefaultValue
        }
        "Yes or No Choice" {
          # Convert the DefaultValue to a boolean
          $qcTaskFormSectionField['ValueAsBoolean'] = `
            [System.Convert]::ToBoolean($qcFormSectionField.DefaultValue)
        }
        { ($_ -eq "Date Only") -or ($_ -eq "Date and Time") } {
          # Convert the DefaultValue to a date
          $qcTaskFormSectionField['ValueAsDate'] = `
            [DateTimeOffset]$qcFormSectionField.DefaultValue
        }
        { ($_ -eq "Single Choice") -or ($_ -eq "Multiple Choice") } {
          # These fields don't have defaults, so they can be ignored unless you want
          # to add selected options during the task creation.
          break;
        }
        default {
          $qcTaskFormSectionField['ValueAsString'] = `
            $qcFormSectionField.DefaultValue
        }
      }
    }

    # Reference fields need to be handled in a different way. They exist as a
    # separate "ReferenceObject" property on the field and contain a list of
    # Values.
    if (($qcFormSectionField.DataTypeName -eq "Reference") `
        -and ($null -ne $qcFormSectionField.ReferenceObject -ne $null)) {

      # Create the reference object property that requires the object id and
      # the list of values.
      $qcTaskFormSectionField['ReferenceObject'] = @{
        ObjectId = $qcFormSectionField.ReferenceObject.ObjectId
        Values   = @()
      }

      # If the form has default reference values, we'll add them here
      $qcFormSectionField.ReferenceObject.DefaultValues | ForEach-Object {
        $refVal = @{
          KeyAsInteger = $_.KeyAsInteger
          KeyAsLong    = $_.KeyAsLong
        }

        # Again, PowerShell will convert $null string properties to an empty string,
        # so we'll handle objects that use a string as an ID separately here.
        if (-not [string]::IsNullOrWhiteSpace($_.KeyAsString)) {
          $refVal["KeyAsString"] = $_.KeyAsString
        }

        # Add the ference values to the reference object
        $qcTaskFormSectionField.ReferenceObject.Values += $refVal
      }
    }

    # Add the field to the section
    $qcTaskFormSection.Fields += $qcTaskFormSectionField
  }
}

# Finally, add the form to the task
$qcTask.Forms += $qcTaskForm

# Uncomment this line to view the task object that will be created
# $qcTask.Forms | ConvertTo-Json -Depth 10

# Use the Agility Blue `Add-Task` PowerShell command followed by the 
# task object to create the task
$createdTask = Add-Task $qcTask

Write-Output "Created QC task $($createdTask.TaskId) in project $($createdTask.ProjectId)"
