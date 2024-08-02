# ------------------------------------------------------------------------------
#
# This script will remove email addresses and associated contacts from an email
# message prior to sending it from the system. It can help with preventing the
# system from sending an email to an unwanted distribution groups that are set
# up to send emails back to Agility Blue
#
# ------------------------------------------------------------------------------
# Event triggers to setup:
#   1) Object: Send Email Message, Action: On Create, Action State: Before Save
#
# ------------------------------------------------------------------------------

# Populate a list of email addresses that you don't want to send emails to
$unwantedEmailAddresses = @(
  "notify@sadiebluesoftware.com"
  "service@sadiebluesoftware.com"
)

function Remove-UnwantedEmailAddressFromString {
  param (
    [string]$emailRecipients,
    [string[]]$unwantedEmailAddresses
  )
    
  # Check if emailRecipients is null or empty
  if ([string]::IsNullOrEmpty($emailRecipients)) {
    return $null
  }
    
  # Split the string into an array
  $emailArray = $emailRecipients -split "; "
    
  # Remove any unwanted email addresses
  $filteredEmailArray = $emailArray | Where-Object { $_ -notin $unwantedEmailAddresses }
    
  # Join the array back into a string
  $filteredEmailRecipients = $filteredEmailArray -join "; "
    
  # Return null if the result is an empty string
  if ([string]::IsNullOrEmpty($filteredEmailRecipients)) {
    return $null
  }
    
  return $filteredEmailRecipients
}

$email = Get-InboundObjectInstance

# We only care about "Outbound" emails
if ($email.Direction -ne "Outbound") {
  Write-Output "Ignoring $($email.Direction) Email"
    
  return $email
}

# These fields are what the notification service uses to send emails
$email.To = Remove-UnwantedEmailAddressFromString `
  -emailRecipients $email.To `
  -unwantedEmailAddresses $unwantedEmailAddresses
    
$email.Cc = Remove-UnwantedEmailAddressFromString `
  -emailRecipients $email.Cc `
  -unwantedEmailAddresses $unwantedEmailAddresses
    
$email.Bcc = Remove-UnwantedEmailAddressFromString `
  -emailRecipients $email.Bcc `
  -unwantedEmailAddresses $unwantedEmailAddresses

# Removing the associated contacts makes it so that unwanted emails
# won't populate into the To/Cc/Bcc fields in the event this email is
# replied to in the UI
$associatedContacts = Get-EmailMessageContacts -Filter "EmailMessageId eq $($email.EmailMessageId)"

$associatedContacts.Collection | ForEach-Object {
  $messageContact = $_
    
  if ($messageContact.ContactEmailAddress -in $unwantedEmailAddresses) {
    Remove-EmailMessageContact -Id $messageContact.EmailMessageContactId | Out-Null
    Write-Output "Removed email message contact $($messageContact.EmailMessageContactId) ($($messageContact.ContactEmailAddress))"
  }
}

# Save the modified email (the send event does not save the email when returned back)
$updatedEmail = Set-EmailMessage -Entry $email

return $updatedEmail