# ------------------------------------------------------------------------------
#
# This is a fun example that showcases how you can call an external API and use
# the response to update a task comment prior to saving it to Agility Blue.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Task Comment, Action: On Create, Action State: Before Save
#
# ------------------------------------------------------------------------------
# For before save event actions to complete successfully, the script must return 
# the same input object type. If anything other than the input object type is returned,
# Agility Blue will treat that as an error and the save will fail. The last output
# of the script can be used to provide the error message to the user.
#
# ------------------------------------------------------------------------------

# The Get-InboundObjectInstance command will give you the object instance that was saved. 
# The type of object is dependent on the trigger context.
$taskComment = Get-InboundObjectInstance

# If the task comment is null just return it back to the pipeline.
if ($null -eq $taskComment) {
  Write-Output "This script needs to be executed by a trigger"
  return $taskComment
}

# The idea here is that we're calling an external API, extracting the result of that
# call, and then appending that result to the incoming task comment. We want to wrap
# the external call in a try-catch block so in the event there was some issue calling
# the service, the task comment will still get saved.
try {
  # The URL to the external service
  $url = "https://api.chucknorris.io/jokes/random"

  # Use PowerShell's Invoke-RestMethod to call the external service.
  $response = Invoke-RestMethod -Uri $url

  # Store an emoji in a variable. This is just for fun.
  $ninjaEmoji = [char]::ConvertFromUtf32(0x1F977)

  # Create a string that will be appended to the task comment. The value property of
  # the response in this case is the quote we'd like to append. We're also adding
  # some HTML markup to make the quote stand out in the task comment.
  $chuckNorrisQuoteToAppend = "<hr /><p><strong>Chuck Norris</strong> quote:</p><blockquote>$ninjaEmoji $($response.value)</blockquote>"

  # Append the quote to the task comment.
  $taskComment.Value += $chuckNorrisQuoteToAppend
} catch {
  # If there was an error calling the external service, we'll just log it to the
  # output stream and continue on. In this case, we don't want the failure of this
  # script to prevent the task comment from being saved.
  Write-Output "Error calling the chuck norris quoting service: $_"
}

# Return the task comment back to Agility Blue. Any properties that were modified
# on the object within the script, it will be saved.
return $taskComment