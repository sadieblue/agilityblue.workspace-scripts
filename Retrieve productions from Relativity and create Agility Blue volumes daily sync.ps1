# ------------------------------------------------------------------------------
#
# This script will reach out to a connected Relativity instance and gather
# productions from linked workspaces. It will then check for matching volumes
# in Agility Blue and create new volumes if none are found (matching, in this
# example, is based on the matter, volume name, and beg/end bates range.
#
# This script is designed to be ran on a daily schedule. It will keep track of
# the last linked workspace id and the date of the last sync in a custom object
# record so that it can pick up where it left off if the script times out and 
# needs to run multiple times in a single day.
# ------------------------------------------------------------------------------
#
# This script can be executed manually or scheduled to run daily.
#
# Scheduled Job Setup:
#   - Minute: 0
#   - Hour: 6
#   - Day: *
#   - Month: *
#   - Weekday: *
#
# (The scheduled job will run every day at 6:00 AM in whatever timezone you set)
#
# If you find that the script is timing out before it can finish, create multiple 
# schedules that run the script 5 minutes apart (for example, 6:00 AM, 6:05 AM, 
# 6:10 AM, etc.). The additional executions will pick up where the last one left
# off.
#
# ------------------------------------------------------------------------------
#
# In order for this script to work properly, a working Relativity integration
# must already be established that has the "Use Modern Relativity Apis" option
# checked within the integration configuration. Check out the help docs if you 
# need to connect to Relativity here: 
# https://help.agilityblue.com/docs/relativity-integration
#
# In addition to a working integration, a custom object must exist that
# helps the script with syncing the data each day. The custom object should 
# have the following characteristics:
#   Name: "Relativity Volumes Sync Data"
#   Fields:
#   1) "Date UTC" (Date Only, Make this field the reference link)
#   2) "Last Linked Relativity Workspace Id" (Whole Number)
#
# Get the id of this custom object and plug it in here:
$relativityVolumesSyncDataCustomObjectId = 10
#
# Other variables that need to be set:
# 
# The time zone to use for dates when creating new volumes
$timeZone = "Central Standard Time"
#
# Your Relativity base URL
relativityUrl = "https://my-company.relativity.one"

# ------------------------------------------------------------------------------

# Some global variables for the final output
$totalLinkedWorkspaces = 0
$matchingVolumes = 0
$createdVolumes = 0

# The time zone we will use for volume dates when creating new volumes
$timeZoneForVolumeDates = [System.TimeZoneInfo]::FindSystemTimeZoneById($timeZone)
$linkedWorkspacesSyncDataDate = (Get-Date).ToString("yyyy-MM-ddT00:00:00Z")

# Retrieve the Relativity integration that we will be using for this script
$integrations = Get-Integrations `
    -Filter "Name eq 'Relativity' and startswith(Identifier,'$relativityUrl')"

$relativityIntegration = $integrations.Collection | Select-Object -First 1

if ($null -eq $relativityIntegration) {
    Write-Error "Unable to find a Relativity integration with the identifier `"$relativityIdentifier`""
    exit
}

# Test the Relativity connection. If Agility Blue doesn't get a valid response from Relativity, terminate 
# the script to reduce further unnecessary calls. The `Test-RelativityConnection` cmdlet returns true
# or false depending on if the connection is successful or not.
$relativityConnectionTestSuccessful = Test-RelativityConnection `
    -IntegrationId $relativityIntegration.IntegrationId

if (-not $relativityConnectionTestSuccessful) {
    Write-Error "Unable to connect to Relativity instance $relativityUrl"
    exit
}

# Get the "Relativity Volumes Sync Data" custom object entry
$relativityVolumesSyncDataObjectSchema = Get-Object -Id $relativityVolumesSyncDataCustomObjectId

$syncDataObjDateField = $relativityVolumesSyncDataObjectSchema.Fields | Where-Object { $_.Label -eq "Date UTC" }
$syncDataObjLastLinkedRelativityWorkspaceIdField = $relativityVolumesSyncDataObjectSchema.Fields | Where-Object { $_.Label -eq "Last Linked Relativity Workspace Id" }

$relativityVolumesSyncData = (Get-CustomObjects -CustomObjectId $relativityVolumesSyncDataCustomObjectId -Top 1).Collection | Select-Object -First 1

# If there are no records in the custom object, create a new one
if (-not $relativityVolumesSyncData) {
    $relativityVolumesSyncData = @{}
    $relativityVolumesSyncData["CF_$($syncDataObjDateField.FieldId)"] = $linkedWorkspacesSyncDataDate
    $relativityVolumesSyncData["CF_$($syncDataObjLastLinkedRelativityWorkspaceIdField.FieldId)"] = $null
    
    $relativityVolumesSyncData = Add-CustomObject `
        -CustomObjectId $relativityVolumesSyncDataCustomObjectId `
        -Entry $relativityVolumesSyncData
}

$syncDataDate = $relativityVolumesSyncData["CF_$($syncDataObjDateField.FieldId)"]
$syncDataObjLastLinkedRelativityWorkspaceId = $relativityVolumesSyncData["CF_$($syncDataObjLastLinkedRelativityWorkspaceIdField.FieldId)"]

# If you want verbose output, create a variable and set it to $true or create a script parameter 
# that you can set when executing the script manually. It's advised not to leave this on for more
# than just troubleshooting as it can create a lot of output.
if ($verboseOutput) {
    Write-Output "Relativity volumes sync data:"
    $relativityVolumesSyncData | ConvertTo-Json -Depth 10
}

$linkedRelativityWorkspacesFilter = "(IntegrationId eq $($relativityIntegration.IntegrationId))"

if ($null -ne $relativityVolumesSyncData["CF_$($syncDataObjLastLinkedRelativityWorkspaceIdField.FieldId)"]) {

    $syncDataDateAsDate = [DateTime]$syncDataDate
    $linkedWorkspacesSyncDataDateAsDate = [DateTime]$linkedWorkspacesSyncDataDate

    if ($verboseOutput) {
        Write-Output "syncDataDate: $syncDataDateAsDate"
        Write-Output "linkedWorkspacesSyncDataDate: $linkedWorkspacesSyncDataDateAsDate"
    }

    if ($syncDataDate -eq $linkedWorkspacesSyncDataDate) {
        # We are still in the middle of gathering data for the current day so create a filter based on the last stored id
        Write-Output "Filtering for linked Relativity workspaces greater than id $syncDataObjLastLinkedRelativityWorkspaceId"
        
        $linkedRelativityWorkspacesFilter += " and (LinkedRelativityWorkspaceId gt $syncDataObjLastLinkedRelativityWorkspaceId)"
    }
}

# Get a list of all Relativity workspaces that are linked to an Agility Blue matter
$linkedRelativityWorkspaces = Get-LinkedRelativityWorkspaces `
    -Filter $linkedRelativityWorkspacesFilter

$totalLinkedWorkspaces = $linkedRelativityWorkspaces.TotalCount

if ($verboseOutput) {
    Write-Output "Linked Relativity workspaces for integration $($relativityIntegration.IntegrationId):"
    $linkedRelativityWorkspaces | ConvertTo-Json -Depth 10
}

# Iterate through each linked Relativity workspace in the collection
$linkedRelativityWorkspaces.Collection | ForEach-Object {
    $linkedRelativityWorkspace = $_
    
    # This call will cause Agility Blue to reach out to Relativity and ask for the productions
    # located in the current linked workspace. The result is an array of Relativity workspace
    # records.
    $relativityWorkspaceProductions = Get-RelativityProductions `
        -IntegrationId $linkedRelativityWorkspace.IntegrationId `
        -WorkspaceArtifactId $linkedRelativityWorkspace.WorkspaceArtifactId
 
    if ($verboseOutput) {
        Write-Output "$($relativityWorkspaceProductions.Count) production(s) in Relativity workspace $($linkedRelativityWorkspace.WorkspaceArtifactId):"
        $relativityWorkspaceProductions | ConvertTo-Json
    }
    
    # Iterate through each Relativity workspace production
    $relativityWorkspaceProductions | ForEach-Object {
        $relativityWorkspaceProduction = $_
        
        # Each Relativity production object has the following fields:
        #   ArtifactID
        #   Name
        #   DateProduced
        #   FirstBatesValue
        #   LastBatesValue
        #   SystemCreatedBy
        #   SystemCreatedOn
        #   Status
        
        # First, we want to make sure that the status is "Produced"
        if ($relativityWorkspaceProduction.Status -ne "Produced") {
            return
        }
        
        # Next, we want to check for a combination of unique fields and match it up with volumes in 
        # the current linked matter in Agility Blue. If your volumes have duplicate names and bates ranges,
        # we'd recommend creating a custom field (such as "Relativity Production Artifact ID") on the volume
        # object to use for more precise volume matching.
        $volumesFilter = "(MatterId eq $($linkedRelativityWorkspace.MatterId))"
        $volumesFilter += " and (Name eq '$($relativityWorkspaceProduction.Name)')"
        $volumesFilter += " and (VolumeRangeBeg eq '$($relativityWorkspaceProduction.FirstBatesValue)')"
        $volumesFilter += " and (VolumeRangeEnd eq '$($relativityWorkspaceProduction.LastBatesValue)')"
        
        if ($verboseOutput) {
            Write-Output "Current volumes filter: $volumesFilter"
        }
        
        # Check if there are any volumes that match the filter. We don't need to get all of the volumes, 
        # only if there are any that match.
        $volumes = Get-Volumes -Filter $volumesFilter -Top 1
        
        if ($verboseOutput) {
            Write-Output "$($volumes.TotalCount) matching volume(s) for Relativity production `"$($relativityWorkspaceProduction.Name)`":"
            $volumes | ConvertTo-Json
        }
        
        if ($volumes.TotalCount -gt 0) {
            $matchingVolumes += $volumes.TotalCount
            
            if ($verboseOutput) {
                Write-Output "There is already an Agility Blue volume named `"$($relativityWorkspaceProduction.Name)`" with a range of $($relativityWorkspaceProduction.FirstBatesValue)-$($relativityWorkspaceProduction.LastBatesValue)"
            }
            
            return
        }
        
        # No volumes in Agility Blue match the filter, so we need to create one
        
        # Relativity provides a DateTime object that we need to convert to a DateTimeOffset because 
        # Agility Blue requires Time zone information for all dates
        $volumeDate = [DateTimeOffset]::new($relativityWorkspaceProduction.DateProduced.Ticks, $timeZoneForVolumeDates.GetUtcOffset($relativityWorkspaceProduction.DateProduced))
            
        $newVolume = @{
            Name        = $relativityWorkspaceProduction.Name
            MatterId    = $linkedRelativityWorkspace.MatterId
            VolumeDate  = $volumeDate
        }
        
        # Wrap the act of creating new volumes within a try/catch block so any errors encountered 
        # here won't break the whole process
        try {
            # Create the new volume in Agility Blue
            $newVolume = Add-Volume -Entry $newVolume
        
            Write-Output "Volume $($newVolume.VolumeId) ($($newVolume.Name)) created on matter $($newVolume.MatterId)"
            
            $createdVolumes++
            
            try {
                $newVolumeRange = @{
                    VolumeId    = $newVolume.VolumeId
                    BegNo       = $relativityWorkspaceProduction.FirstBatesValue
                    EndNo       = $relativityWorkspaceProduction.LastBatesValue
                }
                
                # Create a volume range to add to the volume
                $newVolumeRange = Add-VolumeRange -Entry $newVolumeRange
                
                 Write-Output "Volume range $($newVolumeRange.VolumeRangeId) ($($newVolumeRange.BegNo)-$($newVolumeRange.EndNo)) created on volume $($newVolumeRange.VolumeId) ($($newVolume.Name)) for matter $($newVolume.MatterId)"
            } catch {
                Write-Output "Unable to create volume range $($relativityWorkspaceProduction.FirstBatesValue)-$($relativityWorkspaceProduction.LastBatesValue) on volume $($newVolumeRange.VolumeId) ($($newVolume.Name)) for matter $($linkedRelativityWorkspace.MatterId). Error message: $($_.Exception.Message)"
            
                if ($verboseOutput) {
                    $_.Exception
                }
            }
        } catch {
            Write-Output "Unable to create volume $($relativityWorkspaceProduction.Name) for matter $($linkedRelativityWorkspace.MatterId). Error message: $($_.Exception.Message)"
            
            if ($verboseOutput) {
                $_.Exception
            }
        }
    }
    
    # Update the Relativity volumes sync data custom object record. This will allow the script 
    # to pick up where it left off for another run in the event of a timeout.
    $relativityVolumesSyncData["CF_$($syncDataObjDateField.FieldId)"] = $linkedWorkspacesSyncDataDate
    $relativityVolumesSyncData["CF_$($syncDataObjLastLinkedRelativityWorkspaceIdField.FieldId)"] = $linkedRelativityWorkspace.LinkedRelativityWorkspaceId
    
    $relativityVolumesSyncData = Set-CustomObject `
        -Entry $relativityVolumesSyncData `
        -CustomObjectId $relativityVolumesSyncDataCustomObjectId
}

# Report the results of the script
Write-Output "Out of $($totalLinkedWorkspaces.ToString("N0")) linked Relativity workspace(s), $($matchingVolumes.ToString("N0")) volume(s) already existed and $($createdVolumes.ToString("N0")) new volume(s) were created"
