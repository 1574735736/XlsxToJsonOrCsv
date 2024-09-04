using UnityEngine;
using UnityEditor;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;

public class ExcelToCsvConverter : EditorWindow
{
    private string excelFilePath;
    private string csvFolderPath;
    private bool isConverting;
    //private List<string> excelFilePaths;

    [MenuItem("CustomTools/单个转Excel Single To CSV Converter")]
    public static void ShowWindow()
    {
        ExcelToCsvConverter window = EditorWindow.GetWindow<ExcelToCsvConverter>();
        window.titleContent = new GUIContent("Excel To CSV");
        window.Show();
    }

    private void OnEnable()
    {
        excelFilePath = EditorPrefs.GetString("ExcelFolderPath", "");
        csvFolderPath = EditorPrefs.GetString("CsvFolderPath", "");
    }

    private void OnDestroy()
    {
        EditorPrefs.SetString("ExcelFolderPath", excelFilePath);
        EditorPrefs.SetString("CsvFolderPath", csvFolderPath);
    }

    private void OnGUI()
    {
        GUILayout.Label("选择xlsx文件", EditorStyles.boldLabel);
        excelFilePath = EditorGUILayout.TextField("XLSX 路径:", excelFilePath);
        if (GUILayout.Button("Browse XLSX File", GUILayout.MaxWidth(200)))
        {
            excelFilePath = EditorUtility.OpenFilePanel("Select XLSX File", "", "xlsx");
        }

        GUILayout.Space(10);

        GUILayout.Label("选择输出路径", EditorStyles.boldLabel);
        csvFolderPath = EditorGUILayout.TextField("CSV 路径:", csvFolderPath);
        if (GUILayout.Button("Browse Output Folder", GUILayout.MaxWidth(200)))
        {
            csvFolderPath = EditorUtility.OpenFolderPanel("Select Output Folder", "", "");
        }

        GUILayout.Space(20);

        GUI.enabled = !isConverting;
        if (GUILayout.Button("xlsx转csv", GUILayout.MaxWidth(150)))
        {
            ConvertToCsv();
        }
        GUI.enabled = true;

        GUILayout.Space(20);

        if (GUILayout.Button("打开csv文件夹", GUILayout.MaxWidth(150)))
        {
            OpenCsvFolder();
        }
    }

    private void ConvertToCsv()
    {
        isConverting = true;

        string tempFilePath = Path.GetTempFileName();
        File.Copy(excelFilePath, tempFilePath, true);

        //using (FileStream fileStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
        using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = null;
            try
            {
                workbook = new XSSFWorkbook(fileStream);

            }
            catch (System.Exception)
            {
                workbook = new HSSFWorkbook(fileStream);
            }

            for (int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                string sheetName = sheet.SheetName;
                string csvFileName = Path.GetFileNameWithoutExtension(excelFilePath);

                if (workbook.NumberOfSheets > 1)
                {
                    csvFileName += "_" + sheetName;
                }

                string csvFilePath = Path.Combine(csvFolderPath, csvFileName + ".csv");

                if (File.Exists(csvFilePath))
                {
                    File.Delete(csvFilePath);
                }

                using (StreamWriter streamWriter = new StreamWriter(csvFilePath, false, System.Text.Encoding.UTF8))
                {
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null)
                            continue;

                        for (int cellIndex = 0; cellIndex < row.LastCellNum; cellIndex++)
                        {
                            ICell cell = row.GetCell(cellIndex);
                            string cellValue = "";

                            //string cellValue = (cell == null) ? "" : cell.ToString();
                            if (cell != null)
                            {
                                switch (cell.CellType)
                                {
                                    case CellType.Formula:
                                        switch (cell.CachedFormulaResultType)
                                        {
                                            case CellType.Numeric:
                                                cellValue = cell.NumericCellValue.ToString();
                                                break;
                                            case CellType.String:
                                                cellValue = cell.StringCellValue;
                                                break;
                                        }
                                        break;
                                    default:
                                        cellValue = cell.ToString();
                                        break;
                                }
                            }

                            streamWriter.Write(cellValue);

                            if (cellIndex < row.LastCellNum - 1)
                                streamWriter.Write(",");
                        }
                        streamWriter.WriteLine();
                    }
                }
            }
        }

        File.Delete(tempFilePath);

        isConverting = false;
        EditorUtility.DisplayDialog("Conversion Complete", "XLSX to CSV conversion is complete.", "OK");
    }

   
    private void OpenCsvFolder()
    {
        if (Directory.Exists(csvFolderPath))
        {
            EditorUtility.RevealInFinder(csvFolderPath);
        }
        else
        {
            EditorUtility.DisplayDialog("Folder Not Found", "CSV folder does not exist.", "OK");
        }
    }
}
