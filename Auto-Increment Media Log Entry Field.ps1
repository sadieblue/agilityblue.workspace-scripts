<#

This pre-save script will automatically increment a custom field on the media log entry object on a per matter basis. Example below:

  MatterReference-000001

If the matter does not have a reference, the script will increment the number without the reference and dash as a prefix. Example below:

  000001

Event triggers to setup:
  1) Object: Media Log Entry, Action: On Create, Action State: Before Save

#>

# Get the inbound object from the media log entry
$mediaEntry = Get-InboundObjectInstance

# Gets the MatterId from the inbound object
$matterID = $mediaEntry.MatterId

# Sets the length of the auto-increment number e.g. 000001 (6 digits) - Update as needed
$length = 6

# Set the startVal to the first increment number e.g. 1 = 000001 - Update as needed
$startVal = "1"

# Identifies the field where the auto-increment number will be updated - Update as needed
$targetField = "CF_105"

# Identfies the label value of the custom field - Update as needed
$label = "Serial Number"

# This function will iterate through the returned Media Log Entries collection to find the next available number
Function Get-NextMLEID{
    param ($inputObjects, $matterRef, $length)

    $highestNumber = $inputObjects.Collection | Where-Object {
        # Ensure CF_105 is not null and matches the expected format
        $null -ne $_.$targetField -and $_.$targetField -is [string] -and $_.$targetField -match "-\d{$length}$"
    } | ForEach-Object {
        # Extract the last 6 digits, remove leading zeros, and convert to integer
        [int]($_.$targetField -replace ".*-(\d{$length})", '$1')
    } | Measure-Object -Maximum
    
    $nextNumber = if ($null -ne $highestNumber.Maximum) {
        # Identifies the highest number and increments by 1
        ($highestNumber.Maximum + 1).ToString().PadLeft($length, '0')
    } else {
        # If no highest number is identified, defaults to $startVal and appends zeros identified in $length variable
        $startVal.ToString().PadLeft($length, '0')
    }
    
    # Identifies if the Matter Reference is null or not null
   $nextSerialNumber = if($null -eq $matterRef){
       # If Matter Reference is Null, the format will be 00001
        $nextNumber
    } else{
    # If Matter Reference is not null, the format will be {matter reference}-00001
        $matterRef +"-"+ $nextNumber
    }
    
    # returns the the auto-increment number
    return $nextSerialNumber
}


if ($null -ne $matterID) {
    # Code to execute if the $matterID variable is not null
    Write-Output "Media Log Entry is associated with a Matter. ID will be generated."
    
    # Gets the Matter object
    $matterObj = Get-Matter -Id $matterID
    
    # Sets the matter reference varible for the matter object
    $matterRef = $matterObj.Reference
    
    # Gets the Media Log Entries with MatterId Filter 
    $mle = Get-MediaLogEntries -Top 10 -Filter "MatterId eq $matterID" -OrderBy "MediaLogEntryId desc"
    
    # Runs Get-NextMLEID to find next available number
    $newMLENumber = Get-NextMLEID $mle $matterRef $length
    
    #Displays next auto-increment number
    Write-Output $newMLENumber
    
    # Sets the auto-increment number to CF_105/Serial Number Field
    $serialNumberField = $mediaEntry.Fields | Where-Object { $_.Label -eq "$label" }
    
    if ($null -eq $serialNumberField.Value) {
        $serialNumberField.Value = @{}
        }

    $serialNumberField.Value.ValueAsString = $newMLENumber


} else {
    # Code to execute if the $matterID variable is null
    Write-Output "The media log entry is not associated with a matter. No custom ID will be created."
}

# Returns updated Media Log Entry to complete the save action with updated Serial Number (CF_105) Field.
return $mediaEntry
