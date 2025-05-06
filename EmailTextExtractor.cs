using MimeKit;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

public static class EmailTextExtractor
{
    public static string ExtractText(string filePath)
    {
        try
        {
            using (var stream = File.OpenRead(filePath))
            {
                var message = MimeMessage.Load(stream);
                var text = new StringBuilder();

                // Add email headers
                text.AppendLine($"From: {message.From}");
                text.AppendLine($"To: {message.To}");
                text.AppendLine($"Subject: {message.Subject}");
                text.AppendLine($"Date: {message.Date:g}");
                text.AppendLine();

                // Extract text from the message body
                ExtractPart(message.Body, text);

                return text.ToString();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting text from email {filePath}: {ex.Message}");
            return string.Empty;
        }
    }

    private static void ExtractPart(MimeEntity part, StringBuilder text)
    {
        if (part is TextPart textPart)
        {
            // Handle text content
            if (textPart.IsPlain || textPart.IsRichText)
            {
                text.AppendLine(textPart.Text);
            }
            else if (textPart.IsHtml)
            {
                // Simple HTML to text conversion - just get the text without tags
                var html = textPart.Text;
                if (string.IsNullOrEmpty(html))
                    return;
                    
                var inTag = false;
                foreach (var c in html)
                {
                    if (c == '<')
                    {
                        inTag = true;
                        text.Append(' ');
                    }
                    else if (c == '>')
                    {
                        inTag = false;
                    }
                    else if (!inTag)
                    {
                        text.Append(c);
                    }
                }
                text.AppendLine();
            }
        }
        else if (part is Multipart multipart)
        {
            // Handle multipart content (like multipart/alternative or multipart/mixed)
            foreach (var subpart in multipart)
            {
                ExtractPart(subpart, text);
            }
        }
        else if (part is MessagePart messagePart)
        {
            try
            {
                // Handle message/rfc822 parts (nested messages)
                var message = messagePart.Message;
                if (message != null)
                {
                    text.AppendLine("--- Forwarded Message ---");
                    text.AppendLine($"From: {message.From}");
                    text.AppendLine($"To: {message.To}");
                    text.AppendLine($"Subject: {message.Subject}");
                    text.AppendLine($"Date: {message.Date:g}");
                    text.AppendLine();
                    ExtractPart(message.Body, text);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing message part: {ex.Message}");
            }
        }
    }
}
