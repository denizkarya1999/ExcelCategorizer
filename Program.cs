using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

class Program
{
    // Class-level variables to store the data type and variables
    static string data_type;
    static List<string> variables;
    static List<string> file_paths = new List<string>();

    static void Main()
    {
        string inputFilePath = "C:\\Users\\deniz\\source\\Primary_CAPPS.xlsx";
        string outputFilePath = "C:\\Users\\deniz\\source\\Primary_Processed_CAPPS.xlsx";

        var columnsData = ReadFirstColumnFromFirstWorksheet(inputFilePath);
        var secondColumnsData = ReadSecondColumnFromFirstWorksheet(inputFilePath);
        var processedData = new List<(string, List<string>, string)>();

        for (int i = 0; i < columnsData.Count; i++)
        {
            string columnData = columnsData[i];
            string secondColumnData = secondColumnsData[i];

            // Trigger ParseWords function.
            string[] parsedWords = ParseWords(columnData);

            // Trigger categorize function.
            categorize(parsedWords);

            // Store the processed data with the corresponding file path
            processedData.Add((data_type, variables, secondColumnData));
        }

        // Transform the data into the desired format
        var transformedData = TransformData(processedData);

        // Write the transformed data to a new Excel file
        WriteToExcel(outputFilePath, transformedData);
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
        if (rawCode.Length < 2)
        {
            Console.WriteLine("Invalid data format.");
            return;
        }

        // Initialize the data type.
        data_type = rawCode[1]; // Assuming the first element is the data type.

        // Initialize the variables.
        variables = new List<string>();
        for (int i = 2; i < rawCode.Length; i++)
        {
            // Ensure we don't add the data type itself as a variable
            if (rawCode[i] != data_type)
            {
                variables.Add(rawCode[i]);
            }
        }
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

    // Read the second column.
    static List<string> ReadSecondColumnFromFirstWorksheet(string filePath)
    {
        List<string> secondColumnData = new List<string>();

        // Load the Excel file.
        var fileInfo = new FileInfo(filePath);

        // Set the license context for EPPlus.
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(fileInfo))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets[0];

            // Read the second column starting from the second row.
            int rowCount = worksheet.Dimension.Rows;
            for (int row = 2; row <= rowCount; row++)
            {
                var cellValue = worksheet.Cells[row, 2].Text; // Read from the second column
                secondColumnData.Add(cellValue);
            }
        }

        return secondColumnData;
    }

    // Transform the processed data into the desired format.
    static List<(string dataType, string variable, string filePath)> TransformData(List<(string dataType, List<string> variables, string filePath)> data)
    {
        var transformedData = new List<(string dataType, string variable, string filePath)>();

        foreach (var entry in data)
        {
            foreach (var variable in entry.variables)
            {
                transformedData.Add((entry.dataType, variable, entry.filePath));
            }
        }

        return transformedData;
    }

    // Write the transformed data to a new Excel file.
    static void WriteToExcel(string filePath, List<(string dataType, string variable, string filePath)> data)
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
            worksheet.Cells[1, 2].Value = "Variable";
            worksheet.Cells[1, 3].Value = "File";

            // Write the data
            int row = 2;
            foreach (var entry in data)
            {
                worksheet.Cells[row, 1].Value = entry.dataType;
                worksheet.Cells[row, 2].Value = entry.variable;
                worksheet.Cells[row, 3].Value = entry.filePath;
                row++;
            }

            // Save the package
            package.SaveAs(fileInfo);
        }
    }
}