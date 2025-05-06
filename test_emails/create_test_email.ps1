# Create a new email message
$message = [MimeKit.MimeMessage]::new()
$message.From.Add([MimeKit.MailboxAddress]::Parse("sender@example.com"))
$message.To.Add([MimeKit.MailboxAddress]::Parse("recipient@example.com"))
$message.Subject = "Test Email with Attachment"

# Create the HTML body
$body = [MimeKit.TextPart]::new("html")
$body.Text = @"
<html>
<body>
    <h1>Test Email</h1>
    <p>This is a test email with an attachment.</p>
    <p>It contains some <b>important</b> information.</p>
    <p>Keywords: test, email, attachment, important</p>
</body>
</html>
"@

# Create the multipart/mixed container
var multipart = [MimeKit.Multipart]("mixed")
multipart.Add(body)

# Add the attachment
var attachment = [MimeKit.MimePart]("text", "plain")
attachment.Content = [MimeKit.MimeContent]::new([System.IO.File]::OpenRead("sample_attachment.txt"))
attachment.ContentDisposition = [MimeKit.ContentDisposition]::new([MimeKit.ContentDisposition.Attachment])
attachment.ContentTransferEncoding = [MimeKit.ContentEncoding.Base64]
attachment.FileName = "sample_attachment.txt"

multipart.Add(attachment)

# Set the message body
$message.Body = multipart

# Save the email to a file
$message.WriteTo("test_email.eml")
