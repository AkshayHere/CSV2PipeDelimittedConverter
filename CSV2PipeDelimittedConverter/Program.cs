using ExcelDataReader;
using Serilog;
using System;
using System.Data;
using System.IO;

class CSV2PipeDelimittedConverter
{
    const int MAX_RETRIES = 5;
    const int DELAY_DURATION = 2000;
    const string ROOT_FOLDER = $"C:\\CSVConverter";
    const string FILE_TO_PROCESS = "ZANCINP.xls";

    static void Main(string[] args)
    {
        string dateNow = DateTime.Now.ToString("yyyyMMdd");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.Console()
            .WriteTo.File($"{ROOT_FOLDER}\\Logs\\converter.log", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        string inputFolder = ROOT_FOLDER;
        string outputFolder = $"{ROOT_FOLDER}\\Output";
        string backupFolder = $"{ROOT_FOLDER}\\Backup";

        string inputFile = Path.Combine(inputFolder, FILE_TO_PROCESS);

        // Ensure folders exist
        Directory.CreateDirectory(backupFolder);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Required for ExcelDataReader (for older .xls encoding)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        try
        {
            if (!File.Exists(inputFile))
            {
                Log.Warning($"File not found: {inputFile}");
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputFile = Path.Combine(outputFolder, $"ZANCINP_{timestamp}.txt");

            ConvertExcelToPipe(inputFile, outputFile);

            // Backup the processed file
            Log.Information($"Backing up the processed file.");
            string destPath = Path.Combine(backupFolder, $"ZANCINP_{timestamp}.xls");
            File.Move(inputFile, destPath);

            Log.Information($"File converted successfully: {outputFile}");
        }
        catch (Exception ex)
        {
            Log.Error($"Error: {ex.Message}");
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }

    static void ConvertExcelToPipe(string inputFile, string outputFile)
    {
        try
        {
            if (!File.Exists(inputFile))
            {
                Log.Warning($"Input file does not exist: {inputFile}");
                return;
            }

            using (var stream = OpenFileWithRetry(inputFile))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTable table = result.Tables[0];
                    using (StreamWriter writer = new StreamWriter(outputFile))
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            string[] fields = new string[row.ItemArray.Length];

                            for (int i = 0; i < row.ItemArray.Length; i++)
                            {
                                fields[i] = row[i]?.ToString()?.Trim() ?? "";
                            }

                            string line = string.Join("|", fields);
                            writer.WriteLine(line);
                        }
                    }
                }
            }
        }
        catch (FileNotFoundException ex)
        {
            Log.Error($"File not found: {ex.FileName}");
        }
        catch (Exception ex)
        {
            Log.Error("Unexpected error: " + ex.Message);
        }
    }

    static FileStream OpenFileWithRetry(string path, int maxRetries = MAX_RETRIES, int delayMs = DELAY_DURATION)
    {
        for (int attempt = 1; attempt <= maxRetries; attempt++)
        {
            try
            {
                return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                if (attempt == maxRetries)
                    throw;

                Log.Warning($"File is locked. Retrying {attempt}/{maxRetries}...");
                System.Threading.Thread.Sleep(delayMs);
            }
        }

        throw new Exception("Unable to open file after retries.");
    }
}