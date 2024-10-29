using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI; // Make sure to include this for UI elements
using TMPro;

public class FileManager : MonoBehaviour
{
    public string txtFileType;
    public Button filePickerButton; // Assign this from the inspector
    public TextMeshProUGUI txtHolder;
    public ExcelReaderUnity excelReader;

    void Start()
    {
        /*txtFileType = NativeFilePicker.ConvertExtensionToFileType(".xlsx"); // Handles text file type
*/
        // Ensure the button click triggers the file picking process
        filePickerButton.onClick.AddListener(OnFilePickerButtonClick);
    }

    void OnFilePickerButtonClick()
    {
        // Don't attempt to import/export files if the file picker is already open
        if (NativeFilePicker.IsFilePickerBusy())
            return;

        // Request permission first
        RequestPermissionAsynchronously();
    }

    // Handles the actual file picking and importing after permission
    private void HandleFileOperations()
    {
        // Pick a txt file
        NativeFilePicker.Permission permission = NativeFilePicker.PickFile((path) =>
        {
            if (path == null)
            {
                Debug.Log("Operation cancelled");
            }
            else
            {
                Debug.Log("Picked file: " + path);
                // Read the contents of the selected txt file and display it in txtHolder
                /*ReadTxtFile(path);
*/
                txtHolder.text = excelReader.ReadExcelFile(path);
            }
        });

        Debug.Log("Result Success");
    }

    // Method to read the text file and display its content in txtHolder
    private void ReadTxtFile(string filePath)
    {
        if (File.Exists(filePath))
        {
            string fileContents = File.ReadAllText(filePath); // Read all the content of the text file
            txtHolder.text = fileContents; // Display the text in the txtHolder UI element
            Debug.Log("File contents: " + fileContents);
        }
        else
        {
            Debug.Log("File does not exist at: " + filePath);
        }
    }

    // Request permission asynchronously
    private async void RequestPermissionAsynchronously(bool readPermissionOnly = false)
    {
        NativeFilePicker.Permission permission = await NativeFilePicker.RequestPermissionAsync(readPermissionOnly);
        Debug.Log("Permission result: " + permission);

        if (permission == NativeFilePicker.Permission.Granted)
        {
            // Proceed with file operations if permission is granted
            HandleFileOperations();
        }
        else
        {
            Debug.Log("Permission denied.");
        }
    }
}
