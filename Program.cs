using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

class Program
{
    // Class-level variables to store the data type and variables
    static string data_type;
    static string[] variables;

    static void Main()
    {
        string inputFilePath = "C:\\Primary_CAPPS.xlsx";
        string outputFilePath = "C:\\Users\\deniz\\source\\Primary_Processed_CAPPS.xlsx";

        var columnsData = ReadFirstColumnFromFirstWorksheet(inputFilePath);
        var processedData = new List<(string, string[])>();

        foreach (var columnData in columnsData)
        {
            // Trigger ParseWords function.
            string[] parsedWords = ParseWords(columnData);

            // Trigger categorize function.
            categorize(parsedWords);

            // Store the processed data
            processedData.Add((data_type, variables));
        }

        // Write the processed data to a new Excel file
        WriteToExcel(outputFilePath, processedData);
    }

    // Detect every single word and put them in a string array.
    static string[] ParseWords(string input)
    {
        // Split the input string at the comment delimiter.
        string[] parts = input.Split(new[] { "//" }, StringSplitOptions.None);

        // Use only the part before the comment.
        string beforeComment = parts[0];

        // Use a regular expression to find words.
        MatchCollection matches = Regex.Matches(beforeComment, @"\b\w+\b");

        // A variable to store the words.
        List<string> words = new List<string>();

        // In each Regex match, add the value to the list.
        foreach (Match match in matches)
        {
            words.Add(match.Value);
        }

        // Return the list as an array.
        return words.ToArray();
    }

    // Categorize the parsed string.
    static void categorize(string[] rawCode)
    {
        if (rawCode.Length < 1)
        {
            Console.WriteLine("Invalid data format.");
            return;
        }

        // Initialize the data type.
        data_type = rawCode[1];

        // Initialize the variables.
        List<string> variableList = new List<string>();
        for (int i = 2; i < rawCode.Length; i++)
        {
            variableList.Add(rawCode[i]);
        }

        // Assign the variables to the class-level array.
        variables = variableList.ToArray();
    }

    // Read the excel file.
    static List<string> ReadFirstColumnFromFirstWorksheet(string filePath)
    {
        List<string> firstColumnData = new List<string>();

        // Load the Excel file.
        var fileInfo = new FileInfo(filePath);

        // Set the license context for EPPlus.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(fileInfo))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[0];

            // Read the first column starting from the second row.
            int rowCount = worksheet.Dimension.Rows;
            for (int row = 2; row <= rowCount; row++)
            {
                var cellValue = worksheet.Cells[row, 1].Text;
                firstColumnData.Add(cellValue);
            }
        }

        return firstColumnData;
    }

    // Write the processed data to a new Excel file.
    static void WriteToExcel(string filePath, List<(string dataType, string[] variables)> data)
    {
        // Create a new file info object
        var fileInfo = new FileInfo(filePath);

        // Set the license context for EPPlus.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage())
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Add("Processed Data");

            // Write the headers
            worksheet.Cells[1, 1].Value = "Data Type";

            // Determine the maximum number of variables
            int maxVariables = 0;
            foreach (var entry in data)
            {
                if (entry.variables.Length > maxVariables)
                {
                    maxVariables = entry.variables.Length;
                }
            }

            // Write variable headers dynamically
            for (int i = 1; i <= maxVariables; i++)
            {
                worksheet.Cells[1, i + 1].Value = $"Variable-{i}";
            }

            // Write the data
            int row = 2;
            foreach (var entry in data)
            {
                worksheet.Cells[row, 1].Value = entry.dataType;
                for (int col = 0; col < entry.variables.Length; col++)
                {
                    worksheet.Cells[row, col + 2].Value = entry.variables[col];
                }
                row++;
            }

            // Save the package
            package.SaveAs(fileInfo);
        }
    }
}
