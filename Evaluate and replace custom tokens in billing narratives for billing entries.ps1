<# 

**Billing Entries Narrative token field update with current user**

Within the Billing Entries Form, this script will evaluate the narrative 
field to determine if the string token "#current_user#" is present within 
the string. If the '#current_user#' string token is present, the script 
will replace the token with the current logged in user. This script is 
not exclusive to any single billing narrative, so it will triggered on 
all billing entries.

Event triggers to setup:
  1) Object: Billing Entry, Action: On Create, Action State: Before Save
  2) Object: Billing Entry, Action: On Update, Action State: Before Save

#>

# Retrieve a reference to the inbound billing entry
$billingEntry = Get-InboundObjectInstance

# The $eventTriggeredBy object identifies the logged in user that triggered the event
$eventTriggeredBy = $ABVAREvent.CreatedBy.ToString() | ConvertFrom-Json
    
# The $triggeredByUserName captures the Display Name of the logged in user that triggered the event
$triggeredByUserName = $eventTriggeredBy.DisplayValue

# Sets boolean value if #current_user# text is found in Narrative string value
$match = $billingEntry.Narrative -match "#current_user#"

if ($match) {

  # The '#current_user#' token will be replaced with the $currentuser variable
  $billingEntry.Narrative = $billingEntry.Narrative.Replace("#current_user#", $triggeredByUserName)
    
  Write-Output "Current user token was found, Narrative entry will be updated"
    
}
else {
  Write-Output "Current user token not found"
}

# Returns the updated billing entry back to the system
return $billingEntry