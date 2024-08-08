using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("TableConvertor.exe <input_file> <output_file>");
            return;
        }

        string inputFilePath = args[0];
        string outputFilePath = args[1];

        string inputExtension = Path.GetExtension(inputFilePath).ToLower();
        string outputExtension = Path.GetExtension(outputFilePath).ToLower();

        try
        {
            if (inputExtension == ".csv" && outputExtension == ".xlsx")
            {
                ConvertCsvToXlsx(inputFilePath, outputFilePath);
            }
            else if (inputExtension == ".xlsx" && outputExtension == ".csv")
            {
                ConvertXlsxToCsv(inputFilePath, outputFilePath);
            }
            else
            {
                Console.WriteLine("Unsupported file formats.");
                Console.WriteLine("Supported conversions:");
                Console.WriteLine("CSV to XLSX");
                Console.WriteLine("XLSX to CSV");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }

    static void ConvertCsvToXlsx(string csvFilePath, string xlsxFilePath)
    {
        // 设置 EPPlus 许可上下文
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // 读取 CSV 文件
        List<dynamic> records;
        using (var reader = new StreamReader(csvFilePath))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            records = new List<dynamic>(csv.GetRecords<dynamic>());
        }

        // 写入 XLSX 文件
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            if (records.Count > 0)
            {
                var headers = ((IDictionary<string, object>)records[0]).Keys;
                int colIndex = 1;
                foreach (var header in headers)
                {
                    worksheet.Cells[1, colIndex++].Value = header;
                }

                int rowIndex = 2;
                foreach (var record in records)
                {
                    colIndex = 1;
                    foreach (var value in ((IDictionary<string, object>)record).Values)
                    {
                        worksheet.Cells[rowIndex, colIndex++].Value = value;
                    }
                    rowIndex++;
                }
            }
            package.SaveAs(new FileInfo(xlsxFilePath));
        }

        Console.WriteLine("CSV file has been successfully converted to XLSX.");
    }

    static void ConvertXlsxToCsv(string xlsxFilePath, string csvFilePath)
    {
        // 设置 EPPlus 许可上下文
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // 读取 XLSX 文件
        using (var package = new ExcelPackage(new FileInfo(xlsxFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            using (var writer = new StreamWriter(csvFilePath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                // 写入头部
                for (int col = 1; col <= colCount; col++)
                {
                    csv.WriteField(worksheet.Cells[1, col].Text);
                }
                csv.NextRecord();

                // 写入数据行
                for (int row = 2; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        csv.WriteField(worksheet.Cells[row, col].Text);
                    }
                    csv.NextRecord();
                }
            }
        }

        Console.WriteLine("XLSX file has been successfully converted to CSV.");
    }
}