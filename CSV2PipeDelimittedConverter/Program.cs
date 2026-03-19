using System;
using System.Data;
using System.IO;
using ExcelDataReader;

class CSV2PipeDelimittedConverter
{
    static void Main(string[] args)
    {
        string inputFolder = @"C:\CSVConverter";
        string outputFolder = @"C:\CSVConverter\Output";
        string inputFile = Path.Combine(inputFolder, "ZANCINP.xls");

        // Ensure folders exist
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Required for ExcelDataReader (for older .xls encoding)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        try
        {
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"File not found: {inputFile}");
                return;
            }

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputFile = Path.Combine(outputFolder, $"ZANCINP_{timestamp}.txt");

            ConvertExcelToPipe(inputFile, outputFile);

            Console.WriteLine($"File converted successfully: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static void ConvertExcelToPipe(string inputFile, string outputFile)
    {
        try
        {
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file does not exist: {inputFile}");
                return;
            }

            using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    // Get first sheet
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
            Console.WriteLine($"File not found: {ex.FileName}");
        }
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine("Access denied. Check permissions.");
        }
        catch (IOException ex)
        {
            Console.WriteLine("I/O error occurred: " + ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unexpected error: " + ex.Message);
        }
    }
}