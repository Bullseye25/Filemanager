using System;
using System.IO;
using UnityEngine;
using OfficeOpenXml;  // Ensure you include this namespace for EPPlus

public class ExcelReaderUnity : MonoBehaviour
{
/*    // Set the path to your Excel file relative to the Unity project
    public string filePath = "Assets/YourFolder/yourExcelFile.xlsx";

    void Start()
    {
        // Call the method to read the Excel file
        ReadExcelFile(filePath);
    }*/

    public string ReadExcelFile(string path)
    {
        string result = string.Empty;

        // Verify the file exists
        if (File.Exists(path))
        {
            FileInfo fileInfo = new FileInfo(path);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // Load the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Loop through the rows and columns
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        // Read the cell value
                        var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                        /*Debug.Log($"Cell [{row},{col}]: {cellValue}");
*/
                        if(col == 1)
                            result += $"{cellValue}";
                        else
                            result += $" {cellValue}\n";
                    }
                }
            }
        }
        else
        {
            Debug.LogError("Excel file not found at path: " + path);
        }

        return result;
    }
}
