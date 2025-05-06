using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;
using WordRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;
using WordParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;

public static class OfficeTextExtractor
{
    public static string ExtractText(string filePath)
    {
        try
        {
            var extension = Path.GetExtension(filePath).ToLower();
            
            switch (extension)
            {
                case ".docx":
                    return ExtractTextFromWord(filePath);
                case ".pptx":
                    return ExtractTextFromPowerPoint(filePath);
                case ".xlsx":
                    return ExtractTextFromExcel(filePath);
                default:
                    return string.Empty;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error extracting text from Office document {filePath}: {ex.Message}");
            return string.Empty;
        }
    }

    private static string ExtractTextFromWord(string filePath)
    {
        var text = new StringBuilder();
        
        using (var doc = WordprocessingDocument.Open(filePath, false))
        {
            var body = doc.MainDocumentPart?.Document.Body;
            if (body != null)
            {
                text.Append(GetPlainText(body));
            }
        }
        
        return text.ToString();
    }

    private static string ExtractTextFromPowerPoint(string filePath)
    {
        var text = new StringBuilder();
        
        using (var doc = PresentationDocument.Open(filePath, false))
        {
            var presentationPart = doc.PresentationPart;
            if (presentationPart != null)
            {
                // Get the presentation part
                var slideParts = presentationPart.SlideParts;
                
                // Loop through all the slides and get the text
                foreach (var slidePart in slideParts)
                {
                    if (slidePart?.Slide != null)
                    {
                        // Extract text from all text elements in the slide
                        foreach (var textElement in slidePart.Slide.Descendants<A.Text>())
                        {
                            text.AppendLine(textElement.Text);
                        }
                    }
                }
            }
        }
        
        return text.ToString();
    }

    private static string ExtractTextFromExcel(string filePath)
    {
        var text = new StringBuilder();
        
        using (var doc = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = doc.WorkbookPart;
            if (workbookPart?.Workbook != null)
            {
                var sheets = workbookPart.Workbook.Descendants<Sheet>();
                
                foreach (var sheet in sheets)
                {
                    if (sheet.Id?.Value == null) continue;
                    
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                    var sharedStringTablePart = workbookPart.SharedStringTablePart;
                    
                    if (worksheetPart?.Worksheet != null && sharedStringTablePart?.SharedStringTable != null)
                    {
                        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        if (sheetData == null) continue;
                        
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            foreach (var cell in row.Elements<Cell>())
                            {
                                if (cell.CellValue != null)
                                {
                                    string cellValue = cell.CellValue.Text;
                                    
                                    // Handle shared strings
                                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                                    {
                                        if (int.TryParse(cellValue, out int ssid))
                                        {
                                            var item = sharedStringTablePart.SharedStringTable.ElementAt(ssid);
                                            text.Append(item.InnerText + " ");
                                        }
                                    }
                                    else
                                    {
                                        text.Append(cellValue + " ");
                                    }
                                }
                            }
                            text.AppendLine();
                        }
                    }
                }
            }
        }
        
        return text.ToString();
    }

    private static string GetPlainText(OpenXmlElement element)
    {
        var text = new StringBuilder();
        
        foreach (var paragraph in element.Elements<WordParagraph>())
        {
            foreach (var run in paragraph.Elements<WordRun>())
            {
                foreach (var textElement in run.Elements<WordText>())
                {
                    text.Append(textElement.Text);
                }
                text.Append(" ");
            }
            text.AppendLine();
        }
        
        return text.ToString();
    }
}
