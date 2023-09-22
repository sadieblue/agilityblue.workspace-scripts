# ------------------------------------------------------------------------------
#
# Send a message to Slack when a task is created or completed
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

#region Define Variables

# Place your Agility Blue URL here. It's used by the Slack message to link back to the task.
$myAgilityBlueUrl = "https://agilityblue.com"

# Place your Slack webhook URL here. You get this from the Slack.
$slackWebhookUrl = "https://hooks.slack.com/services/YOUR_SLACK_WEBHOOK_URL"

#endregion

#region Initialization ---------------------------------------------------------

$task = $agilityBlueObject

if ($null -eq $task) {
  # If the object is not available, it means we are executing the script 
  # manually. In this case, we'll get a task directly for testing purposes.
  $task = Get-Task 1000
  
  Write-Output "### Using test task $($task.TaskId) ###"
}

# Uncomment this line to view the available properties/values
# $task | ConvertTo-JSON -Depth 10

#endregion

# We only care about normal projects (ignore drafts)
if ($task.ProjectTypeName -ne "Normal") {
  Write-Output "Task $($task.TaskId) does not belong to a normal project. Skipping further processing."
}

$slackHeader = "New Agility Blue Task"
$slackEmoji = ":large_blue_circle:"

# Check the event object to see if this is an update action. If so, we are only interested in completed tasks.
if ($agilityBlueEvent.Action -eq 2 -and $task.StatusDescription -ne "Completed") {
  Write-Output "Task updated, but not completed. Skipping further processing."
} 

if ($task.StatusDescription -eq "Completed") {
  $slackHeader = "Agility Blue Task Completed"
  $slackEmoji = ":white_check_mark:"
}

$due = "Ongoing"

if ($null -ne $task.DateDue) {
  $due = "Due on $([DateTimeOffset]::Parse($task.DateDue).ToString("g"))"
}

# Slack expects a particular object structure to create a message. Create that here.
$slackMessage = [ordered]@{
  blocks      = @(
    [ordered]@{
      type = "header"
      text = [ordered]@{
        type  = "plain_text"
        text  = $slackHeader
        emoji = $true
      }
    }
    [ordered]@{
      type = "section"
      text = [ordered]@{
        type = "mrkdwn"
        text = "$($slackEmoji) Task #$($task.TaskId): *$($task.Name)*"
      } 
    }
    [ordered]@{
      type = "section"
      text = [ordered]@{
        type = "mrkdwn"
        text = "*Client*: $($task.ClientName)`n*Matter*: $($task.MatterName)`n`n<$($myAgilityBlueUrl)/workspace/$($agilityBlueEvent.WorkspaceId)/tasks/$($task.TaskId)|View Task in Agility Blue>"
      }
    }
  )
  attachments = @(
    [ordered]@{
      blocks = @(
        [ordered]@{
          type = "section"
          text = [ordered]@{
            type = "mrkdwn"
            text = "Created by $($task.CreatedByFullName)`n$($due)"
          }
        }
      )
    }
  )
}

# If the task is completed, add the completed by information to the message.
if ($task.StatusDescription -eq "Completed") {
  $slackMessage.attachments[0].blocks += [ordered]@{
    type = "section"
    text = [ordered]@{
      type = "mrkdwn"
      text = "Completed by $($task.CompletedByFullName)"
    }
  }
}

# Convert our slack message to JSON and send it to Slack using PowerShell's Invoke-RestMethod cmdlet.
$slackMessageAsJson = $slackMessage | ConvertTo-Json -Depth 10 -Compress

Invoke-RestMethod -Uri $slackWebhookUrl -Method Post -Body $slackMessageAsJson -ContentType "application/json" | Out-Null

Write-Output "Slack message sent for task $($task.TaskId)"
