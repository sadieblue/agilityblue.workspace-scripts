# ------------------------------------------------------------------------------
#
# Checks that the matter reference field is 6 characters long. If it's not, the
# script will prevent the matter from being saved. The end user will be notified
# that the script prevented the save.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Matter, Action: On Create, Action State: Before Save
#   1) Object: Matter, Action: On Update, Action State: Before Save
#
# ------------------------------------------------------------------------------
# For before save event actions to complete successfully, the script must return 
# the same input object type. If anything other than the input object type is returned,
# Agility Blue will treat that as an error and the save will fail. The last output
# of the script can be used to provide the error message to the user.
#
# ------------------------------------------------------------------------------

# The `Get-InboundObjectInstance` command will provide scripts with the incoming
# object instance based on the trigger object. In this case, the trigger object
# is a matter.
$matter = Get-InboundObjectInstance

if ($null -eq $matter) {
  # If the inbound object instance is null, it means the script is being executed
  # manually. In this case, we'll retrieve a specific matter for testing purposes.
  $matter = Get-Matter -Id 1
}

if (($null -eq $matter.Reference) -or ($matter.Reference.Length -ne 6)) {
  # By returning something other than the input object type (a matter in this case), 
  # Agility Blue will treat this as an error and prevent the save.
  return "The matter reference field must be 6 characters long"
}

# Return the matter back to Agility Blue to allow the save to continue.
return $matter
