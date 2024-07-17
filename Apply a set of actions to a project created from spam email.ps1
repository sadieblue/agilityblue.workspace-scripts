# ------------------------------------------------------------------------------
#
# If you receive inbound email that is created as projects within Agility Blue,
# you may run into situations where the email is spam. You may want to keep these
# projects around instead of deleting them for analysis on the type of spam you
# are receiving to aide in creating tighter spam filters. To save time and 
# pontential error-prone user mistakes, this script will tag a provided project 
# id as spam, create comments on any tasks within that project that cancel the 
# tasks, move the project to a matter that's used to contain these spam projects, 
# and finally cancel the project.
#
# ------------------------------------------------------------------------------
# [Manual Execution]
#
# Add a required Basic Text parameter with a label of "Project Id" and a name of 
# "ProjectId"
# ------------------------------------------------------------------------------
# In order for this script to work properly, you need to provide a tag id and a 
# matter id below in the USER-DEFINED VARIABLES section. We recommend a tag named
# "Spam" and an internal matter named "Spam Emails"
#
# ------------------------------------------------------------------------------
# USER-DEFINED VARIABLES

# The id of the tag that will be applied to the project. Tags are added as a list
# of tag ids to a project, so we wrap the id inside an array.
$spamTagIdToAdd     = @(92)

# The id of the matter that the project will be moved to
$spamMatterId       = 150

# ------------------------------------------------------------------------------

# Define a function here that we'll use to make sure that the project id parameter
# provided by the user is a number
function Is-Number {
    param (
        [string]$inputString
    )

    $regex = '^\d+$'
    
    return $inputString -match $regex
}

# Test the ProjectId paramter
if (-not (Is-Number -inputString $ProjectId)) {
    Write-Error "$ProjectId is not a number"
    exit
}

# ProjectId is a number, so continue with retrieving the project
$project = Get-Project -Id $ProjectId

if ($null -eq $project) {
    Write-Output "Unable to locate a project with an id of $ProjectId"
    exit
}

# Tag the project as spam
# -----------------------------------------------------------------------------------------
if (($null -ne $project.ComputedTags) -and ($project.ComputedTags.Contains("Spam"))) {
    Write-Output "Project $($project.ProjectId) is already tagged as spam"
} else {
    Add-TagsToProject `
        -Id $project.ProjectId `
        -TagIds $spamTagIdToAdd `
        -Append
        
    Write-Output "Project $($project.ProjectId) tagged as spam"
}

# Get a list of tasks within the project and cancel them by creating cancel comments
# -----------------------------------------------------------------------------------------
$tasks = Get-Tasks -Filter "ProjectId eq $($project.ProjectId)"

Write-Output "Iterating over $($tasks.TotalCount) task(s) in project $($project.ProjectId)..."

$tasks.Collection | ForEach-Object {
    $task = $_
    
    # To cancel a task, we need to create a cancel comment (type 4)
    $taskComment = @{
        TaskId          = $task.TaskId
        CommentTypeId   = 4
        Value           = "Spam content"
    }
    
    Add-TaskComment -Entry $taskComment | Out-Null
    
    Write-Output "Canceled task $($task.TaskId) on project $($project.ProjectId)"
}

# Convert the project to a normal project, set the matter to the spam matter and cancel it
# -----------------------------------------------------------------------------------------
$project.MatterId       = $spamMatterId
$project.ProjectTypeId  = 1
$project.IsCanceled     = $true

$updatedProject = Set-Project -Entry $project

Write-Output "Canceled project $($updatedProject.ProjectId)"
