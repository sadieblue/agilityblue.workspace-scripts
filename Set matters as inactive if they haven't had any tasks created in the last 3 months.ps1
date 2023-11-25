# ------------------------------------------------------------------------------
#
# This script will locate all matters that haven't had any tasks created
# in them over the last 3 months and set them to inactive.
#
# ------------------------------------------------------------------------------
# [Manual Execution]
#
# ------------------------------------------------------------------------------

# Use the `ConvertTo-DateString` command to convert a DateTime object that represents
# the start of the day in central time to a string format that Agility Blue expects 
# for filters.
$dateString = ConvertTo-DateString `
  -DateTime (Get-Date).Date.AddDays(-90) `
  -TimeZone "Central Standard Time" `
  -IncludeTime

# Define a filter to find active matters where the last project created was longer
# than 90 days ago
$filter = "(Active eq true) and (LastProjectCreatedOnDate lt $dateString)"

# Retreive the matters using the filter
$matters = Get-Matters -Filter $filter -Top 1000

if ($matters.TotalCount -eq 0) {
  # No projects were found, exit the script
  Write-Output "No active matters found where the last project created was earlier than $dateString"
  Exit
}

Write-Output "Found $($matters.TotalCount) active matters where the last project created was earlier than $dateString"

$mattersUpdated = 0

do {
  # Iterate through each matter and set it to inactive
  $matters.Collection | ForEach-Object {
    $matterInCollection = $_

    try {
      # The matters resulting in the `Get-Matters` can't be used to update 
      # modify matters because the type of model the system exepcts for 
      # creating/updating isn't the same. Therefore, we need to retrieve each 
      # matter we want to update individually using the `Get-Matter` command 
      # and then update it.
      $matter = Get-Matter -Id $matterInCollection.MatterId

      $matter.Active = $false

      # The `Set-Matter` command will update the matter in Agility Blue
      # and return the updated matter. We don't actually need the returned 
      # matter in this case, so we'll pipe it to Out-Null to discard it.
      Set-Matter -Entry $matter | Out-Null

      $mattersUpdated++
    } catch {
      # Something went wrong with updating a specific matter, so we'll log
      # the error and continue processing the rest of the matters.
      Write-Output "Failed to set matter $($matter.MatterId) to inactive: $_"
    }
  }

  # Re-run the filter to see if there are any more records to process
  $matters = Get-Matters -Filter $filter -Top 1000

} while ($matters.TotalCount -gt 0)

Write-Output "Updated $mattersUpdated matters"
