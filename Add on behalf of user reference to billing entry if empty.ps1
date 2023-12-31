# ------------------------------------------------------------------------------
#
# If a billing entry is saved without an on behalf of user, this script will add the
# created by user to the on behalf of user field.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Billing Entry, Action: On Create, Action State: After Save
#   2) Object: Billing Entry, Action: On Update, Action State: After Save
#
# ------------------------------------------------------------------------------
# In order for this script to work properly, the Billing Entry object must have a user 
#   reference field named "On Behalf Of"
#
# ------------------------------------------------------------------------------

# Retrieve the inbound id of the billing entry that was created
$billingEntryId = Get-InboundObjectId

if ($null -eq $billingEntryId) {
  # If the id is null, it probably means we are executing the script manually.
  # In this case, we'll set an id directly for testing purposes
  $billingEntryId = 6
}

# Note here that we want to get a full billing entry that includes all custom fields and
# references. The injected Get-InboundObjectInstance will only provide a shallow version.
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
  $billingEntry = Set-BillingEntry $billingEntry -BypassCustomFieldValidation
  
  Write-Output "The on behalf of user field for billing entry $($billingEntry.BillingEntryId) has been updated"
} else {
  Write-Output "The on behalf of user field for billing entry $($billingEntry.BillingEntryId) already has a value"
}
