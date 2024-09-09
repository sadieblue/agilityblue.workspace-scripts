# ------------------------------------------------------------------------------
#
# This script will create a comment on any tasks that are late, which sends out
# a notification to users that have notification rules setup for the task. The
# script will only create one comment per day per task to avoid spamming users
# with multiple notifications for the same task.
#
# ------------------------------------------------------------------------------
#
# The scheduled job will run every hour, 5 minutes after the hour, between 9:00 AM 
# and 5:00 PM
#
# Scheduled Job Setup:
#   - Minute: 5
#   - Hour: 9-17
#   - Day: *
#   - Month: *
#   - Weekday: MON-FRI
#
# ------------------------------------------------------------------------------

# Create a message that will be used as our late task comment.
$lateTaskMessage = @"
<p>
    This task is late. Please take care of it!
</p>
<hr />
<p>
    This message was sent automatically by a scheduled job
</p>
"@.Trim()

# Get the current date and time in a format that Agility Blue likes in a specific time zone so it's not
# using the stored UTC times for our filters
$currentDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([System.DateTimeOffset]::Now, "Central Standard Time")
$todayDateTime = [System.DateTimeOffset]::new($currentDateTime.Date, $currentDateTime.Offset)

# Get our dates that will be used for filtering
$now = ConvertTo-DateString -DateTimeOffset $currentDateTime -IncludeTime
$today = ConvertTo-DateString -DateTimeOffset $todayDateTime -IncludeTime

# Retreive a collection of tasks that are late
$lateTasks = Get-Tasks -Filter "(StatusDescription ne 'Completed') and (StatusDescription ne 'Canceled') and (DateDue le $now)"

$lateTasks.Collection | ForEach-Object {
    $lateTask = $_
    
    Write-Output "Task $($lateTask.TaskId) is $($lateTask.StatusDescription) and was due on $($lateTask.DateDue)"
    
    # Create a filter that we will use to get any comments that were already created today about the job being late so we don't
    # send the same message multiple times in a day.
    $taskCommentsFilter = "contains(Value,'this task is late') and (TaskId eq $($lateTask.TaskId)) and (CommentTypeName eq 'Issue') and (CreatedOn ge $today)"
    
    # Get any comments that were created today about the job being late. We want our filter to be unique enough
    # here to catch only comments that were created by the automation
    $taskComments = Get-TaskComments -Filter $taskCommentsFilter
    
    if ($taskComments.TotalCount -eq 0) {
        # Create a new task comment on the task
        # A comment type of 2 adds an "Issue" badge to the comment within the UI
        $taskComment = @{
            TaskId = $lateTask.TaskId
            CommentTypeId = 2
            Value = $lateTaskMessage
        }
        
        $newComment = Add-TaskComment -Entry $taskComment
        
        # We need to tell the system to invoke a notification so people that have notification rules setup
        # for this task will get a notification from Agility Blue with our message
        # More info about invoking notifications, see https://help.agilityblue.com/docs/scripting-invoke-notification
        Invoke-Notification -RequestType "New Task Comment" -Id $newComment.TaskCommentId
        
        Write-Output "Comment $($newComment.TaskCommentId) created on task $($newComment.TaskId) because it is late"
    } else {
        Write-Output "A late task comment for task $($lateTask.TaskId) has already been created for today"
    }
}