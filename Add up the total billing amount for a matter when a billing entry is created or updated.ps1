# ------------------------------------------------------------------------------
#
# When a user creates or updates a billing entry, this script will add up all of 
# the billing entry costs for the matter and update the total billing amount on a 
# custom field on the matter.
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Billing Entry, Action: On Create, Action State: After Save
#   2) Object: Billing Entry, Action: On Update, Action State: After Save
#
# ------------------------------------------------------------------------------
# Custom Fields to setup:
#   1) Object: Matter, Name: Total Billing Amount, Type: Decimal
#
# ------------------------------------------------------------------------------
# Caveats:
#   1) This script does not handle deleted billing entries because by the time the
#      script is executed, the billing entry is already deleted and cannot be queried.
#      This can be handled by converting this script to a pre-save event script.  
#   2) This script does not handle billing entries that have been moved to a 
#      different matter where the billing entry was not directly updated (such as
#      moving the project or task the billing entry is attached to). To accomodate 
#      these scenarios, you would need different scripts to handle those events.
#
# ------------------------------------------------------------------------------

# This culture info is just used to display the currency using a dollar symbol in the log
$usCulture = [Globalization.CultureInfo]::GetCultureInfo("en-US")

# The Get-InboundObjectInstance command will give you the object instance that was saved. 
# The type of object is dependent on the trigger context.
$billingEntry = Get-InboundObjectInstance

if ($null -eq $billingEntry) {
  # If the event object is not available, it means we are executing the script 
  # manually. In this case, we'll get a billingEntry directly for testing purposes.
  $billingEntry = Get-BillingEntry 88950
  
  Write-Output "### Using test billing entry $($billingEntry.BillingEntryId) ###"
}

# Uncomment this line if you want to view the available properties/values
# $billingEntry | ConvertTo-JSON

# First, we want to make sure that the billing entry has an associated matter
if ($null -eq $billingEntry.MatterId) {
  Write-Output "Billing entry does not have an associated matter"
  Exit
}

# Next, get all billing entries for the matter. We'll want to implement a paging strategy here because
# the API can only return 1000 records at a time. We'll use the Top parameter to get the first 1000 records,
# and then use the Skip parameter to get the next 1000 records, and so on until we get all of the records.

$sum = 0
$count = 0

$billingEntriesTop = 1000
$billingEntriesSkip = 0

$billingEntriesFilter = "MatterId eq $($billingEntry.MatterId)"

do {
  $matterBillingEntries = Get-BillingEntries `
    -Filter $billingEntriesFilter `
    -Top $billingEntriesTop `
    -Skip $billingEntriesSkip

  $count += $matterBillingEntries.Collection.Count

  # Calculate the sum of the billing entries. Here, we can use powershell's Measure-Object cmdlet
  $sum += ($matterBillingEntries.Collection | Measure-Object -Property { $_.Quantity * $_.UnitPrice } -Sum).Sum

  $billingEntriesSkip += $billingEntriesTop
} while ($billingEntriesSkip -lt $matterBillingEntries.TotalCount)

# Finally, update the matter with the total billing amount
$matter = Get-Matter $billingEntry.MatterId

# Locate the custom field on the matter named "Total Billing Amount" and data type of "Decimal"
$totalBillingAmountCustomField = $matter.Fields | Where-Object { $_.Label -eq "Total Billing Amount" -and $_.DataTypeName -eq "Decimal Number" }

if ($null -eq $totalBillingAmountCustomField) {
  Write-Output "Decimal custom field 'Total Billing Amount' not found on matter $($matter.MatterId)"

  # Uncomment this line to help troubleshoot why the field may be reported as missing
  # $matter | ConvertTo-Json -Depth 10
  Exit
}

# Update the custom field with the sum of the billing entries
$totalBillingAmountCustomField.Value = @{
  ValueAsDecimal = $sum
}

# Save the matter
$matter = Set-Matter $matter

Write-Output "Matter '$($matter.Name)' (id: $($matter.MatterId)) contains $($count.ToString('N0')) billing entries totaling $($sum.ToString('C', $usCulture))"