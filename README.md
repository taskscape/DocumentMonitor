# Document Monitor and Search Application

This is a C# console application that monitors specified folders for document changes and allows you to search through the documents using Lucene.NET.

## Features

- Monitors specified folders for document changes (additions and modifications)
- Supports multiple document formats:
  - Word Documents (.docx)
  - PowerPoint Presentations (.pptx)
  - Excel Spreadsheets (.xlsx)
  - PDF Documents (.pdf)
  - Markdown Files (.md)
  - Text Files (.txt)
  - Email Messages (.eml)
- Full-text search with Lucene.NET
- Real-time indexing of new and modified documents
- Search results ranked by relevance

## Prerequisites

- .NET 9.0 SDK or later
- Windows OS (for file system monitoring)

## Configuration

1. Edit the `appsettings.json` file to specify which folders to monitor:
   ```json
   {
     "MonitoredFolders": [
       "C:\\Path\\To\\Monitor",
       "D:\\Another\\Folder\\To\\Monitor"
     ]
   }
   ```

   If no folders are specified, it will default to monitoring a "Monitor" folder in your Documents directory.

## How to Use

1. Build the application:
   ```
   dotnet build
   ```

2. Run the application:
   ```
   dotnet run
   ```

3. The application will:
   - Create the Lucene index if it doesn't exist
   - Start monitoring the specified folders
   - Index any existing documents in those folders

4. While the application is running:
   - Press `S` to search for documents
   - Press `Q` to quit the application

## Search Syntax

The search supports Lucene query syntax. Some examples:

- Simple search: `keyword`
- Phrase search: `"exact phrase"`
- Wildcard: `test*`
- Boolean operators: `AND`, `OR`, `NOT`
- Field-specific search: `filename:report` or `content:important`

## How It Works

1. The application uses `FileSystemWatcher` to monitor the specified folders for file changes.
2. When a document is added or modified, it's processed and indexed by Lucene.
3. The search functionality uses Lucene's query parser to find matching documents.
4. Results are displayed with the filename, modification date, and relevance score.

## Notes

- The Lucene index is stored in a `LuceneIndex` folder in the application directory.
- The application must have read access to the monitored folders.
- For best performance, avoid monitoring folders with a large number of files or very large files.
