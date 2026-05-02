# Create an Outlook application COM object
$outlook = New-Object -ComObject Outlook.Application

# Get the namespace
$namespace = $outlook.GetNamespace("MAPI")

# Access the Inbox (you can change this to another folder if needed)
$inbox = $namespace.GetDefaultFolder(6) # 6 refers to the Inbox folder

# Check if there are any items in the Inbox
if ($inbox.Items.Count -eq 0) {
    Write-Host "The Inbox is empty."
} else {
    # Get the first email in the Inbox
    $email = $inbox.Items.GetFirst()

    # Check if the email is not null
    if ($email -ne $null) {
        # Display the subject of the email
        Write-Host "Subject: $($email.Subject)"

        # Get the RFC headers
        $headers = $email.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")

        # Display the headers
        Write-Host "RFC Headers:"
        Write-Host $headers
    } else {
        Write-Host "No email found."
    }
}
