# ----------------------------------------------------------------------------------------
#
# In order for this script to work properly, the Billing Entry object must have a user 
# reference field named "On Behalf Of"
#
# ----------------------------------------------------------------------------------------

if ($null -eq $agilityBlueEvent) {
  # If the event object is not available, it means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $billingEntryId = 6
} else {
  # The script was executed by an event, so we'll have access to the event object
  $billingEntryId = $agilityBlueEvent.Payload.Id
}

# Note here that we want to get a full billing entry that includes all custom fields and
# references. The injected $agilityBlueObject will only provide a shallow version.
$billingEntry = Get-BillingEntry $billingEntryId

# Uncomment this line to view the available properties/values
# $billingEntry | ConvertTo-JSON -Depth 10

# Get the "On Behalf Of" user reference field
$onBehalfOfField = $billingEntry.Fields | Where-Object { $_.Label -eq "On Behalf Of" }

# A field can be null if the field was added after the billing entry was created
if ($null -eq $onBehalfOfField) {
    $onBehalfOfField = @{}
}

if ($null -eq $onBehalfOfField.Value) {
    $onBehalfOfField.Value = @{}
}

if ($null -eq $onBehalfOfField.Value.ReferenceObject) {
    $onBehalfOfField.Value.ReferenceObject = @{
      Values = @()
    }
}

# Check if the field contains any values. If it does, do nothing.
if ($onBehalfOfField.Value.ReferenceObject.Values.Count -eq 0) {
  # There are no reference entries, add the created by user
  $onBehalfOfField.Value.ReferenceObject.Values += @{ KeyAsString = $billingEntry.CreatedById }
  
  # Save the billing entry
  $billingEntry = Set-BillingEntry $billingEntry -BypassCustomFieldValidation $true
  
  Write-Output "The on behalf of user field for billing entry $($billingEntry.BillingEntryId) has been updated"
} else {
  Write-Output "The on behalf of user field for billing entry $($billingEntry.BillingEntryId) already has a value"
}
