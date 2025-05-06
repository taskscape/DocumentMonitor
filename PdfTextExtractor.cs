using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.IO;
using System.Text;

public static class PdfTextExtractor
{
    public static string ExtractText(string filePath)
    {
        try
        {
            var pdfReader = new PdfReader(filePath);
            var pdfDocument = new PdfDocument(pdfReader);
            var text = new StringBuilder();
            
            for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
            {
                var page = pdfDocument.GetPage(i);
                var strategy = new SimpleTextExtractionStrategy();
                var currentText = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(page, strategy);
                text.AppendLine(currentText);
            }
            
            pdfDocument.Close();
            pdfReader.Close();
            
            return text.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting text from PDF {filePath}: {ex.Message}");
            return string.Empty;
        }
    }
}
