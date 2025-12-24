# ------------------------------------------------------------------------------
#
# This script is used to automatically link Relativity One workspaces to existing
# Agility Blue matters.
#
# When executed, the script will:
#   • Retrieve all enabled and valid Relativity One integrations
#   • Enumerate the Relativity workspaces associated with each integration
#   • Skip system, template, maintenance, and automation workspaces based on
#     configurable name patterns
#   • Check whether each remaining workspace is already linked to an Agility Blue
#     matter
#   • Attempt to find a matching Agility Blue matter based on an exact name match
#   • Create a link between the Relativity workspace and the matched matter when
#     a match is found
#
# This script is designed to be safely re-run. Workspaces that are already linked
# will be detected and skipped to prevent duplicate links from being created.
#
# ------------------------------------------------------------------------------
# [Matching Behavior]
#
# By default, matching is performed using an exact comparison between the
# Relativity workspace name and the Agility Blue matter name. This behavior can
# be customized in the Find-RelativityWorkspaceMatter function if your naming
# conventions require more flexible matching logic.
#
# ------------------------------------------------------------------------------
# [Customization]
#
# Workspace name patterns that should be ignored (such as templates or system
# workspaces) are defined in the USER-DEFINED VARIABLES section below and can be
# modified to suit your environment.
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

# Finds a matter for a provided Relativity workspace
function Find-RelativityWorkspaceMatter {
    [CmdletBinding()]
    param (
        $relativityWorkspace
    )
    
    # Matching strategy:
    #   In this example, we are simply matching the Relativity workspace name to an Agility Blue matter name
    #   and returning the matter if matched.
    
    # Clean the workspace name so any names with quotes don't break the odata filter
    $workspaceName = $relativityWorkspace.Name.Replace("'", "''")

    $matching = Get-Matters `
        -Filter "Name eq '$workspaceName'" `
        -Top 1
    
    if ($matching.TotalCount -gt 0) {
        return $matching.Collection[0]
    } else {
        return $null
    }
}

# Processes each workspace for a provided Relativity integration
function Process-RelativityWorkspaces {
    [CmdletBinding()]
    param (
        $integration
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
            Write-Output "Relativity workspace $($relativityWorkspace.Name) is already linked"
            continue
        }
    
        # No link exists, check if there is a matching matter
        $matchingMatter = Find-RelativityWorkspaceMatter -relativityWorkspace $relativityWorkspace
    
        if ($null -ne $matchingMatter) {
            Write-Output "Matched '$($relativityWorkspace.Name)' with '$($matchingMatter.Name)'. Creating a link..."
            
            $link = @{
                IntegrationId = $integration.IntegrationId
                MatterId = $matchingMatter.MatterId
                WorkspaceArtifactId = $relativityWorkspace.ArtifactID
                WorkspaceName = $relativityWorkspace.Name
            }
            
            $createdLink = Add-LinkedRelativityWorkspace -Entry $link
            
            Write-Output "Created link id #$($createdLink.LinkedRelativityWorkspaceId) between Relativity workspace '$($createdLink.WorkspaceName)' and Agility Blue matter '$($createdLink.MatterName)'"
        } else {
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
    
    Process-RelativityWorkspaces -integration $integration
}
