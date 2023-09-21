# ------------------------------------------------------------------------------
#
# When ran, this script will copy a grid state from one user to all other users.
# If a user already has a state with the same name, they will be skipped.
#
# ------------------------------------------------------------------------------
# [Manual Execution]
#
# ------------------------------------------------------------------------------

$gridId = "mattersGrid"
$gridStateName = "Active Matters"
$sourceUserId = "37004c9a-93e3-473c-8f96-a0556bf6735c"

# ------------------------------------------------------------------------------

# Define the filters to find the grid state to copy
$gridStateFilters = @(
  @{ Field = "GridId"; Value = $gridId }
  @{ Field = "Name"; Value = $gridStateName }
  @{ Field = "UserId"; Value = $sourceUserId }
)

$gridStates = Get-GridStates -Filter $gridStateFilters

if ($gridStates.TotalCount -eq 0) {
  Write-Output "No grid state found to copy"
  Exit
}

$gridStateToCopy = $gridStates.Collection[0]

# Define the filters to find the users to copy the grid state to
# For example, it only makes sense to copy a grid state to users that
# are active, not deleted, not system accounts, and not a team account.
# Note that if "Type" is not defined, it defaults to "Where", and if
# "Operator" is not defined, it defaults to "=".
$workspaceUsersFilter = @(
  @{ Field = "Active"; Value = $true }
  @{ Field = "Deleted"; Value = $false }
  @{ Field = "IsSystemAccount"; Value = $false }
  @{ Field = "IsTeamAccount"; Value = $false }
  @{ Field = "Id"; Operator = "<>"; Value = $sourceUserId }
)

# Get the collection of server-filtered workspace users
$usersToCopyTo = Get-WorkspaceUsers -Filters $workspaceUsersFilter -Top 1000

if ($usersToCopyTo.TotalCount -eq 0) {
  Write-Output "No users found"
  Exit
}

# Filter the list down further within the script to only admin/orgnaization users...
$usersToCopyTo = $usersToCopyTo.Collection | Where-Object {
  $_.Roles -like "*Organization Administrator*" -or $_.Roles -like "*Organization User*"
}

if ($usersToCopyTo.TotalCount -eq 0) {
  Write-Output "No users found"
  Exit
}

# Iterate through each user and first check to see if they already have the grid state,
# if not, then copy it to them.
$usersToCopyTo | ForEach-Object {
  $user = $_
  
  $gridStateFilters = @(
    @{ Field = "GridId"; Value = $gridId }
    @{ Field = "Name"; Value = $gridStateName }
    @{ Field = "UserId"; Value = $user.Id }
  )

  $gridStates = Get-GridStates -Filter $gridStateFilters

  # We only want to create the grid state if the grid states call returns 0 results.
  if ($gridStates.TotalCount -eq 0) {
    # Update the UserId parameter to the current user and create the grid state for them.
    $gridStateToCopy.UserId = $user.Id

    Add-GridState $gridStateToCopy | Out-Null # Out-Null to suppress the output

    Write-Output "Copied grid state '$($gridStateToCopy.Name)' to $($user.FullName)"
  }
  else {
    Write-Output "$($user.FullName) already has grid state '$($gridStateToCopy.Name)'"
  }
}
