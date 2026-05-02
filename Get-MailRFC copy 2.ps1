# Example: Search for an email with a specific subject
$subjectToFind = "Добро пожаловать в систему электронной почты MDaemon для домена vezu.ru"
$email = $inbox.Items | Where-Object { $_.Subject -eq $subjectToFind } | Select-Object -First 1

if ($email -ne $null) {
    # Display the subject of the email
    Write-Host "Subject: $($email.Subject)"

    # Get the RFC headers
    $headers = $email.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")

    # Display the headers
    Write-Host "RFC Headers:"
    Write-Host $headers
} else {
    Write-Host "No email found with the specified subject."
}
