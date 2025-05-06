using MimeKit;
using MimeKit.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public static class EmailTextExtractor
{
    public static string ExtractText(string filePath)
    {
        try
        {
            using var stream = File.OpenRead(filePath);
            var message = MimeMessage.Load(stream);
            var text = new StringBuilder();

            // Add subject
            if (!string.IsNullOrEmpty(message.Subject))
            {
                text.AppendLine($"Subject: {message.Subject}");
            }

            // Add from
            if (message.From.Count > 0)
            {
                text.AppendLine($"From: {message.From}");
            }

            // Add to
            if (message.To.Count > 0)
            {
                text.AppendLine($"To: {message.To}");
            }

            // Add date
            if (message.Date != default)
            {
                text.AppendLine($"Date: {message.Date}");
            }

            // Add body text
            text.AppendLine(GetMessageText(message.Body));

            // Add attachment content
            var attachmentText = ExtractAttachments(message).Result;
            if (!string.IsNullOrEmpty(attachmentText))
            {
                text.AppendLine("\n--- ATTACHMENTS ---\n");
                text.AppendLine(attachmentText);
            }

            return text.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting text from email {filePath}: {ex.Message}");
            return string.Empty;
        }
    }

    private static string GetMessageText(MimeEntity entity)
    {
        if (entity is TextPart textPart)
        {
            // If it's HTML, we'll extract just the text
            if (textPart.IsHtml)
            {
                var html = textPart.Text;
                return HtmlToPlainText(html);
            }
            return textPart.Text;
        }
        else if (entity is Multipart multipart)
        {
            var text = new StringBuilder();
            foreach (var part in multipart)
            {
                text.AppendLine(GetMessageText(part));
            }
            return text.ToString();
        }
        return string.Empty;
    }

    private static async Task<string> ExtractAttachments(MimeMessage message)
    {
        var text = new StringBuilder();
        var attachments = new List<MimePart>();
        var supportedExtensions = new[] { ".pdf", ".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt" };

        foreach (var attachment in message.Attachments)
        {
            if (attachment is MimePart part)
            {
                var fileName = part.FileName ?? "unknown";
                var extension = Path.GetExtension(fileName).ToLowerInvariant();

                if (supportedExtensions.Contains(extension))
                {
                    try
                    {
                        using var memory = new MemoryStream();
                        await part.Content.DecodeToAsync(memory);
                        memory.Position = 0;

                        string attachmentText = string.Empty;
                        var tempFilePath = Path.Combine(Path.GetTempPath(), fileName);

                        try
                        {
                            // Save the attachment to a temporary file
                            await using (var fileStream = File.Create(tempFilePath))
                            {
                                await memory.CopyToAsync(fileStream);
                            }

                            // Extract text based on file type
                            switch (extension)
                            {
                                case ".pdf":
                                    attachmentText = PdfTextExtractor.ExtractText(tempFilePath);
                                    break;
                                case ".docx":
                                case ".xlsx":
                                case ".pptx":
                                case ".doc":
                                case ".xls":
                                case ".ppt":
                                    attachmentText = OfficeTextExtractor.ExtractText(tempFilePath);
                                    break;
                            }
                        }
                        finally
                        {
                            // Clean up the temporary file
                            if (File.Exists(tempFilePath))
                            {
                                try { File.Delete(tempFilePath); } catch { /* Ignore cleanup errors */ }
                            }
                        }

                        if (!string.IsNullOrEmpty(attachmentText))
                        {
                            text.AppendLine($"[Attachment: {fileName}]\n{attachmentText}\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing attachment {fileName}: {ex.Message}");
                    }
                }
            }
        }

        return text.ToString();
    }

    private static string HtmlToPlainText(string html)
    {
        if (string.IsNullOrEmpty(html))
            return string.Empty;

        var text = new StringBuilder();
        var inTag = false;
        var inScript = false;
        var inStyle = false;
        var lastChar = '\0';

        for (int i = 0; i < html.Length; i++)
        {
            char c = html[i];

            if (c == '<')
            {
                inTag = true;
                // Check if this is a script or style tag
                if (i + 7 < html.Length && html.Substring(i, 7).Equals("<script", StringComparison.OrdinalIgnoreCase))
                    inScript = true;
                else if (i + 6 < html.Length && html.Substring(i, 6).Equals("<style", StringComparison.OrdinalIgnoreCase))
                    inStyle = true;
                continue;
            }
            else if (c == '>')
            {
                inTag = false;
                // Add a space after certain tags for better readability
                if (!inScript && !inStyle && lastChar != ' ')
                    text.Append(' ');
                lastChar = ' ';
                continue;
            }

            // Skip content inside script or style tags
            if (inScript || inStyle)
            {
                string endTag = html.Substring(i, Math.Min(9, html.Length - i)).ToLowerInvariant();
                if (inScript && endTag.StartsWith("</script>"))
                {
                    inScript = false;
                    i += 8; // Skip past the closing tag
                }
                else if (inStyle && endTag.StartsWith("</style>"))
                {
                    inStyle = false;
                    i += 7; // Skip past the closing tag
                }
                continue;
            }

            if (!inTag)
            {
                // Convert HTML entities and special characters
                if (c == '&')
                {
                    int end = html.IndexOf(';', i);
                    if (end > i)
                    {
                        string entity = html.Substring(i + 1, end - i - 1);
                        switch (entity)
                        {
                            case "amp": text.Append('&'); break;
                            case "lt": text.Append('<'); break;
                            case "gt": text.Append('>'); break;
                            case "quot": text.Append('"'); break;
                            case "nbsp": text.Append(' '); break;
                            // Add more HTML entities as needed
                            default: text.Append(' '); break;
                        }
                        i = end;
                        lastChar = ' ';
                        continue;
                    }
                }

                // Normalize whitespace
                if (char.IsWhiteSpace(c))
                {
                    if (lastChar != ' ' && lastChar != '\n' && lastChar != '\r')
                    {
                        text.Append(' ');
                        lastChar = ' ';
                    }
                }
                else
                {
                    text.Append(c);
                    lastChar = c;
                }
            }
        }

        return text.ToString();
    }

}