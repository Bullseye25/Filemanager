using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Collections.Generic;
using UnityEngine;

public class ExcelParser : MonoBehaviour
{
    public void Execute(string path)
    {
        // Call the method to parse the Excel file into a text format
        string outputPath = Path.Combine(Application.persistentDataPath, "output.txt");
        ParseXlsxToText(path, outputPath);
    }

    void ParseXlsxToText(string xlsxPath, string outputPath)
    {
        // Create a temp directory to extract the Excel contents
        string extractPath = Path.Combine(Application.temporaryCachePath, "xlsx_extract");
        if (Directory.Exists(extractPath))
        {
            Directory.Delete(extractPath, true); // Clean up before extracting
        }

        Directory.CreateDirectory(extractPath);

        // Step 1: Unzip the .xlsx file
        try
        {
            ZipFile.ExtractToDirectory(xlsxPath, extractPath);
            Debug.Log("Extracted .xlsx file");
        }
        catch (Exception ex)
        {
            Debug.LogError("Failed to extract .xlsx file: " + ex.Message);
            return;
        }

        // Step 2: Load shared strings (if any)
        string sharedStringsPath = Path.Combine(extractPath, "xl/sharedStrings.xml");
        List<string> sharedStrings = new List<string>();

        if (File.Exists(sharedStringsPath))
        {
            XmlDocument sharedStringsXml = new XmlDocument();
            sharedStringsXml.Load(sharedStringsPath);
            XmlNodeList stringNodes = sharedStringsXml.GetElementsByTagName("si");

            foreach (XmlNode stringNode in stringNodes)
            {
                sharedStrings.Add(stringNode.InnerText);
            }
        }

        // Step 3: Parse the sheet1.xml
        string sheetPath = Path.Combine(extractPath, "xl/worksheets/sheet1.xml");
        if (!File.Exists(sheetPath))
        {
            Debug.LogError("sheet1.xml not found in the extracted .xlsx");
            return;
        }

        XmlDocument sheetXml = new XmlDocument();
        sheetXml.Load(sheetPath);

        using (StreamWriter writer = new StreamWriter(outputPath))
        {
            XmlNodeList rowNodes = sheetXml.GetElementsByTagName("row");

            // Step 4: Loop through each row and cell
            foreach (XmlNode rowNode in rowNodes)
            {
                string rowText = "";
                XmlNodeList cellNodes = rowNode.ChildNodes;

                foreach (XmlNode cellNode in cellNodes)
                {
                    string cellType = cellNode.Attributes["t"]?.Value;
                    XmlNode cellValueNode = cellNode.SelectSingleNode("v");

                    if (cellValueNode != null)
                    {
                        string cellValue = cellValueNode.InnerText;

                        // Handle shared string reference
                        if (cellType == "s" && int.TryParse(cellValue, out int sharedStringIndex))
                        {
                            rowText += sharedStrings[sharedStringIndex] + "\t";
                        }
                        else
                        {
                            // Direct value (number or inline string)
                            rowText += cellValue + "\t";
                        }
                    }
                    else
                    {
                        rowText += "\t"; // Empty cell
                    }
                }
                writer.WriteLine(rowText.TrimEnd('\t')); // Write row to file
            }
        }

        // Clean up: Delete the extracted files
        Directory.Delete(extractPath, true);

        Debug.Log("Data extracted to text file at: " + outputPath);
    }
}
