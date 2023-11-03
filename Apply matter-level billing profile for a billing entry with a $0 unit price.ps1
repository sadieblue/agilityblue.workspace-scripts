# ------------------------------------------------------------------------------
#
# If a billing entry is saved where the unit price is $0, this script will check
# the matter level billing profile.
#
# The purpose of this script is aimed at importing billing entries either through
# the API or through the SFTP where the unit price is not provided. The script
# can set the unit price based on the matter level billing profile so the user
# doesn't have to explicitly provide the unit price.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Billing Entry, Action: On Create, Action State: After Save
#
# For pre-save event actions to complete successfully, the script must return 
# the same input object type. If anything other than the input object type is returned,
# Agility Blue will treat that as an error and the save will fail. The last output
# of the script can be used to provide the error message to the user.
# ------------------------------------------------------------------------------

# The Get-InboundObjectInstance command will give you the object instance that was saved. 
# The type of object is dependent on the trigger context.
$billingEntry = Get-InboundObjectInstance

if ($null -eq $billingEntry) {
  # If the agility blue object is not available, it means we are executing the script 
  # manually. In this case, we'll set the billing entry directly for testing purposes.
  $billingEntry = Get-BillingEntry 461
}

# Uncomment this line to view the available billing entry properties/values
# $billingEntry | ConvertTo-JSON -Depth 10

# Check that the billing entry has a matter id and the unit price is set to 0
if ($billingEntry.MatterId -gt 0 -and $billingEntry.UnitPrice -eq 0) {
  Write-Output "The unit price for billing entry $($billingEntry.BillingEntryId) is 0.00, checking the matter level billing profile..."

  # The matter level billing profile can be retrieved using the following powershell cmdlet
  $billingProfile = Get-MatterBillingProfile $billingEntry.MatterId -Top 1000
  
  # Uncomment this line to output the billing profile
  # $billingProfile | ConvertTo-JSON -Depth 10

  # Get the billing type from the matter profile
  $billingType = $billingProfile.Collection | Where-Object { $_.BillingTypeId -eq $billingEntry.BillingTypeId }

  if ($null -eq $billingType) {
    Write-Output "Unable to find billing type $($billingEntry.BillingTypeId) in the matter level billing profile for billing entry $($billingEntry.BillingEntryId)"
  } else {
    # Set the unit price to the standard unit price if the unit price is not set.
    # The unit price is only set if it's been overriden at the matter level.
    if ($null -eq $billingType.UnitPrice) {
      $billingEntry.UnitPrice = $billingType.StandardUnitPrice
    } else {
      $billingEntry.UnitPrice = $billingType.UnitPrice
    }

    # Only save the billing entry if the script was executed manually to simulate
    # a pre-save event action
    if ($null -eq $agilityBlueObject) {
      Set-BillingEntry $billingEntry -BypassCustomFieldValidation | Out-Null
      Write-Output "Billing entry $($billingEntry.BillingEntryId) has been directly saved"
    }

    Write-Output "The unit price for billing entry $($billingEntry.BillingEntryId) has been updated to $($billingEntry.UnitPrice)"
  }
}

return $billingEntry