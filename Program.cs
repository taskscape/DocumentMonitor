using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Microsoft.Extensions.Configuration;
using Lucene.Net.Analysis.Standard;
using Lucene.Net.Documents;
using Lucene.Net.Index;
using Lucene.Net.QueryParsers.Classic;
using Lucene.Net.Search;
using Lucene.Net.Store;
using Lucene.Net.Util;
using LuceneVersion = Lucene.Net.Util.LuceneVersion;
using Document = Lucene.Net.Documents.Document;
using System.Linq;
using LuceneDirectory = Lucene.Net.Store.Directory;

class Program
{
    private static readonly string _indexPath = Path.Combine(Environment.CurrentDirectory, "LuceneIndex");
    private static LuceneDirectory? _directory;
    private static IndexWriter? _writer;
    private static readonly List<string> _monitoredFolders = new();
    private static bool _isRunning = true;
    private const LuceneVersion AppLuceneVersion = LuceneVersion.LUCENE_48;

    static void Main(string[] args)
    {
        Console.WriteLine("Document Monitor and Search Application");
        Console.WriteLine("------------------------------------");

        // Initialize Lucene
        InitializeLucene();

        // Load configuration
        var configuration = new ConfigurationBuilder()
            .SetBasePath(System.IO.Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .Build();
            
        // Get monitored folders from configuration or use default
        var folders = configuration.GetSection("MonitoredFolders").GetChildren()
            .Where(x => !string.IsNullOrEmpty(x.Value))
            .Select(x => x.Value!)
            .ToList();
            
        if (folders.Count == 0)
        {
            folders = new List<string> { System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Monitor") };
        }
        _monitoredFolders.AddRange(folders);
        
        // Ensure the monitored folders exist
        foreach (var folder in _monitoredFolders)
        {
            if (!System.IO.Directory.Exists(folder))
            {
                System.IO.Directory.CreateDirectory(folder);
                Console.WriteLine($"Created monitored folder: {folder}");
            }
        }

        Console.WriteLine("\nMonitored folders:");
        foreach (var folder in _monitoredFolders)
        {
            Console.WriteLine($"- {folder}");
        }

        // Start monitoring
        Console.WriteLine("\nStarting file system watchers...");
        var watchers = new List<FileSystemWatcher>();
        foreach (var folder in _monitoredFolders)
        {
            var watcher = new FileSystemWatcher(folder);
            watcher.Created += OnFileCreated;
            watcher.Changed += OnFileChanged;
            watcher.EnableRaisingEvents = true;
            watchers.Add(watcher);
            Console.WriteLine($"Watching: {folder}");
        }

        // Initial index of existing files
        Console.WriteLine("\nIndexing existing files...");
        foreach (var folder in _monitoredFolders)
        {
            IndexExistingFiles(folder);
        }

        Console.WriteLine("\nDocument Monitor is running. Press 'S' to search, 'Q' to quit.");

        // Main loop
        while (_isRunning)
        {
            if (Console.KeyAvailable)
            {
                var key = Console.ReadKey(true);
                if (key.Key == ConsoleKey.S)
                {
                    Console.Write("\nEnter search query: ");
                    var query = Console.ReadLine();
                    if (!string.IsNullOrWhiteSpace(query))
                    {
                        SearchDocuments(query);
                    }
                }
                else if (key.Key == ConsoleKey.Q)
                {
                    _isRunning = false;
                }
            }
            Thread.Sleep(100);
        }

        // Cleanup
        foreach (var watcher in watchers)
        {
            watcher.Dispose();
        }
        _writer?.Dispose();
        _directory?.Dispose();

        Console.WriteLine("\nDocument Monitor has been stopped.");
    }

    private static void InitializeLucene()
    {
        // Create Lucene index directory if it doesn't exist
        if (!System.IO.Directory.Exists(_indexPath))
        {
            System.IO.Directory.CreateDirectory(_indexPath);
        }
        var dirInfo = new System.IO.DirectoryInfo(_indexPath);
        _directory = Lucene.Net.Store.FSDirectory.Open(dirInfo);
        var analyzer = new StandardAnalyzer(AppLuceneVersion);
        var config = new IndexWriterConfig(AppLuceneVersion, analyzer)
        {
            OpenMode = OpenMode.CREATE_OR_APPEND
        };
        _writer = new IndexWriter(_directory, config);
    }

    private static void IndexExistingFiles(string folderPath)
    {
        try
        {
            var extensions = new[] { ".docx", ".pptx", ".xlsx", ".pdf", ".md", ".txt", ".eml" };
            
            foreach (var file in System.IO.Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories))
            {
                if (extensions.Contains(System.IO.Path.GetExtension(file).ToLower()))
                {
                    IndexFile(file);
                }
            }
            _writer?.Commit();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error indexing files: {ex.Message}");
        }
    }

    private static void IndexFile(string filePath)
    {
        try
        {
            var fileInfo = new FileInfo(filePath);
            var extension = Path.GetExtension(filePath).ToLower();
            string content = string.Empty;

            // Extract text based on file type
            switch (extension)
            {
                case ".txt":
                case ".md":
                    content = File.ReadAllText(filePath);
                    break;
                case ".pdf":
                    content = PdfTextExtractor.ExtractText(filePath);
                    break;
                case ".docx":
                case ".pptx":
                case ".xlsx":
                    content = OfficeTextExtractor.ExtractText(filePath);
                    break;
                case ".eml":
                    content = EmailTextExtractor.ExtractText(filePath);
                    break;
            }

            // Create or update document in index
            var searchQuery = new TermQuery(new Term("path", filePath));
            _writer?.DeleteDocuments(searchQuery);

            var doc = new Document
            {
                new TextField("content", ExtractTextFromFile(filePath), Field.Store.YES),
                new StringField("path", filePath, Field.Store.YES),
                new StringField("filename", fileInfo.Name, Field.Store.YES),
                new StringField("extension", fileInfo.Extension.ToLower(), Field.Store.YES),
                new StringField("modified", 
                    DateTools.TimeToString(
                        (long)(fileInfo.LastWriteTimeUtc - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds,
                        DateResolution.MILLISECOND), 
                    Field.Store.YES)
            };
            
            _writer?.AddDocument(doc);
            _writer?.Commit();
            Console.WriteLine($"Indexed: {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error indexing {filePath}: {ex.Message}");
        }
    }

    private static string ExtractTextFromFile(string filePath)
    {
        var fileInfo = new FileInfo(filePath);
        var extension = Path.GetExtension(filePath).ToLower();
        string content = string.Empty;

        // Extract text based on file type
        switch (extension)
        {
            case ".txt":
            case ".md":
                content = File.ReadAllText(filePath);
                break;
            case ".pdf":
                content = PdfTextExtractor.ExtractText(filePath);
                break;
            case ".docx":
            case ".pptx":
            case ".xlsx":
                content = OfficeTextExtractor.ExtractText(filePath);
                break;
            case ".eml":
                content = EmailTextExtractor.ExtractText(filePath);
                break;
        }

        return content;
    }

    private static void SearchDocuments(string queryText)
    {
        if (_writer == null)
        {
            Console.WriteLine("Error: Index writer is not initialized.");
            return;
        }

        try
        {
            using (var reader = _writer.GetReader(false))
            {
                var searcher = new IndexSearcher(reader);
                var analyzer = new StandardAnalyzer(AppLuceneVersion);
                
                var parser = new MultiFieldQueryParser(
                    AppLuceneVersion,
                    new[] { "filename", "content" },
                    analyzer);
                
                var query = parser.Parse(queryText);
                var hits = Array.Empty<ScoreDoc>();
                
                if (query != null && searcher != null)
                {
                    var searchResults = searcher.Search(query, null, 20, Sort.RELEVANCE);
                    hits = searchResults?.ScoreDocs ?? Array.Empty<ScoreDoc>();
                }

                Console.WriteLine($"\nFound {hits.Length} results:");
                if (searcher != null)
                {
                    foreach (var hit in hits)
                    {
                        var foundDoc = searcher.Doc(hit.Doc);
                        if (foundDoc != null)
                        {
                            var modifiedDate = foundDoc.Get("modified");
                            var modified = !string.IsNullOrEmpty(modifiedDate) ? DateTools.StringToDate(modifiedDate).ToString() : "Unknown";
                            var filename = foundDoc.Get("filename") ?? "Unknown";
                            Console.WriteLine($"- {filename} (Modified: {modified}) - Score: {hit.Score:F2}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error searching: {ex.Message}");
        }
    }

    private static void OnFileCreated(object sender, FileSystemEventArgs e)
    {
        if (e.ChangeType == WatcherChangeTypes.Created || e.ChangeType == WatcherChangeTypes.Changed)
        {
            IndexFile(e.FullPath);
            _writer?.Commit();
        }
    }

    private static void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        // Small delay to ensure the file is no longer locked
        Thread.Sleep(500);
        OnFileCreated(sender, e);
    }
}
