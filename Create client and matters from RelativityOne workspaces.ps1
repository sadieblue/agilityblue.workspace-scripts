# ------------------------------------------------------------------------------
# ******************************************************************************
#
# NOTICE: THIS SCRIPT WILL ONLY WORK IF YOU HAVE RELATIVITYONE INTEGRATIONS SET UP
# IN YOUR AGILITY BLUE ENVIRONMENT AND IF YOU HAVE METRICS COLLECTION ENABLED FOR
# THOSE INTEGRATIONS.
#
# ******************************************************************************
#
# This script is used to automatically create clients and matters from Relativity 
# One workspaces.
#
# When executed, the script will:
#   - Retrieve all enabled and valid RelativityOne integrations
#   - Enumerate the Relativity workspaces associated with each integration
#   - Skip system, template, maintenance, and automation workspaces based on
#     configurable name patterns
#   - Check whether each remaining workspace is already linked to an Agility Blue
#     matter
#   - If the workspace is not linked, create a new client (if needed) and matter
#     (if needed) based on RelOne collected metrics
#   - Create a link between the Relativity workspace and the matter
#
# This script is designed to be safely re-run and would recommend running it on a
# daily job schedule.
#
# CAVEAT: IF A WORKSPACE DOES NOT HAVE ANY METRICS COLLECTED YET BECAUSE IT WAS
# JUST CREATED, IT WILL BE SKIPPED UNTIL RELATIVITY COLLECTS METRICS FOR IT,
# TYPICALLY THE FOLLOWING DAY.
#
# Workspaces that are already linked to an Agilitywill be detected 
# and skipped to prevent unnecessary processing.
#
# ------------------------------------------------------------------------------
# [Customization]
#
# Workspace name patterns that should be ignored (such as templates or system
# workspaces) are defined in the USER-DEFINED VARIABLES section below and can be
# modified to suit your environment.
#
# You may have additional fields to populate when creating clients and matters.
#
# ------------------------------------------------------------------------------
# USER-DEFINED VARIABLES

# Workspace name patterns that should be ignored (templates, system workspaces, etc.)
$IgnoredWorkspaceNamePatterns = @(
  '^\[DO NOT ACCESS\]',
  'Template',
  'Maintenance',
  'Automation'
)

# ------------------------------------------------------------------------------
# FUNCTIONS

function Is-RelativityWorkspaceAlreadyLinked {
  [CmdletBinding()]
  param(
    $relativityWorkspace,
    $integration
  )
    
  $linked = Get-LinkedRelativityWorkspaces `
    -Filter "IntegrationId eq $($integration.IntegrationId) and WorkspaceArtifactId eq $($relativityWorkspace.ArtifactID)" `
    -Top 1
    
  return $linked.TotalCount -gt 0
}

# Processes each workspace for a provided Relativity integration
function Process-RelativityWorkspaces {
  [CmdletBinding()]
  param (
    $integration,
    [int]$clientMetricFieldId,
    [int]$matterMetricFieldId
  )
    
  $relativityWorkspaces = Get-RelativityWorkspaces -IntegrationId $integration.IntegrationId
    
  foreach ($relativityWorkspace in $relativityWorkspaces) {
    # First, check if a Relativity workspace name contains an ignored name pattern
    $shouldSkip = $false
        
    foreach ($pattern in $IgnoredWorkspaceNamePatterns) {
      if ($relativityWorkspace.Name -match $pattern) {
        $shouldSkip = $true
        break
      }
    }
        
    if ($shouldSkip) {
      Write-Output "Skipping non-matter workspace '$($relativityWorkspace.Name)'"
      continue
    }
    
    # Next, check if there is already a link for the Relativity workspace and short-circuit if there is
    if (Is-RelativityWorkspaceAlreadyLinked -relativityWorkspace $relativityWorkspace -integration $integration) {
      Write-Output "Relativity workspace '$($relativityWorkspace.Name)' ($($relativityWorkspace.ArtifactID)) is already linked"
      continue
    }
    
    # No link exists, retreive the client and matter names from the Relativity workspace metrics
    $recentMetrics = Get-RelativityMonthlyMetrics `
      -Filter "WorkspaceArtifactId eq $($relativityWorkspace.ArtifactID) and RelativityMetricFieldId in ($clientMetricFieldId,$matterMetricFieldId)" `
      -OrderBy "LastMetricTimestamp desc" `
      -Top 2

    # Pull the two IDs (and validate they exist)
    $clientName = ($recentMetrics.Collection |
      Where-Object FieldName -eq 'Client Name' |
      Select-Object -First 1 -ExpandProperty ValueAsText
    )

    $matterName = ($recentMetrics.Collection |
      Where-Object FieldName -eq 'Matter Name' |
      Select-Object -First 1 -ExpandProperty ValueAsText
    )

    if (-not $clientName -or -not $matterName) {
      Write-Error "Missing required metric fields for workspace $($relativityWorkspace.Name). Expected 'Client Name' and 'Matter Name'."
      continue
    }

    Write-Output "Checking Relativity workpace '$($relativityWorkspace.Name)' ($($relativityWorkspace.ArtifactID)) against client '$clientName' and matter '$matterName'..."

    # Check if a matter exists for the client/matter combination; if not, create them
    $matter = (Get-Matters -Filter "Name eq '$matterName' and ClientName eq '$clientName'" -Top 1).Collection | Select-Object -First 1

    if ($null -eq $matter) {
      # Create the client if it doesn't already exist
      $client = (Get-Clients -Filter "Name eq '$clientName'" -Top 1).Collection | Select-Object -First 1

      if ($null -eq $client) {
        # NOTE: You may have additional fields to populate here based on your needs
        $clientEntry = @{
          Name = $clientName
        }

        $client = Add-Client -Entry $clientEntry

        Write-Output "Created client '$($client.Name)' with id #$($client.ClientId)"
      }

      # NOTE: You may have additional fields to populate here based on your needs
      $matterEntry = @{
        Name     = $matterName
        ClientId = $client.ClientId
      }

      $matter = Add-Matter -Entry $matterEntry
      
      Write-Output "Created matter '$($matter.Name)' with id #$($matter.MatterId)"
    }

    # Sanity check - ensure we have a matter before creating the link
    if ($null -ne $matter) {
      $link = @{
        IntegrationId       = $integration.IntegrationId
        MatterId            = $matter.MatterId
        WorkspaceArtifactId = $relativityWorkspace.ArtifactID
        WorkspaceName       = $relativityWorkspace.Name
      }
            
      $createdLink = Add-LinkedRelativityWorkspace -Entry $link
            
      Write-Output "Created link id #$($createdLink.LinkedRelativityWorkspaceId) between Relativity workspace '$($createdLink.WorkspaceName)' ($($createdLink.WorkspaceArtifactId)) and Agility Blue matter '$($createdLink.MatterName)' (#$($createdLink.MatterId))"
    }
    else {
      Write-Output "No match for '$($relativityWorkspace.Name)'"
    }
  }
}

# ------------------------------------------------------------------------------
# MAIN ENTRY POINT

# Retrieve each Relativity One integration
$integrations = Get-Integrations -Filter "Name eq 'Relativity' and contains(Identifier,'relativity.one') and Enabled eq true and ConfigurationIsValid eq true"

foreach ($integration in $integrations.Collection) {
  Write-Output "Processing $($integration.Identifier)"
    
  # Before we process any workspaces, ensure that we have any RelativityOne metrics collected. These are needed
  # because they contain information about the client and matter associated with each workspace.
  $metricFields = Get-RelativityMetricFields `
    -Filter "IntegrationId eq $($integration.IntegrationId) and Name in ('Client Name','Matter Name') and Enabled eq true"

  # Pull the two IDs (and validate they exist)
  $clientMetricFieldId = ($metricFields.Collection |
    Where-Object Name -eq 'Client Name' |
    Select-Object -First 1 -ExpandProperty RelativityMetricFieldId
  )

  $matterMetricFieldId = ($metricFields.Collection |
    Where-Object Name -eq 'Matter Name' |
    Select-Object -First 1 -ExpandProperty RelativityMetricFieldId
  )

  if (-not $clientMetricFieldId -or -not $matterMetricFieldId) {
    Write-Error "Missing required metric fields for IntegrationId $($integration.IntegrationId). Expected 'Client Name' and 'Matter Name'."
    continue
  }

  Process-RelativityWorkspaces `
    -integration $integration `
    -clientMetricFieldId $clientMetricFieldId `
    -matterMetricFieldId $matterMetricFieldId
}
