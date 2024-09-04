using System.IO;
using UnityEditor;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using System;
using System.Diagnostics;
using Debug = UnityEngine.Debug;

public class XlsxToCsv : EditorWindow
{
    string inputPath = "";
    string outputPath = "";
    int lineCount = 0;
    //List<string> shieldNames;

    [MenuItem("XlsxTools/通用Xlsx to Csv")]
    public static void ShowWindow()
    {
        EditorWindow.GetWindow(typeof(XlsxToCsv));
    }

    Dictionary<string, string> csvFileOuts = new Dictionary<string, string>();

    private void OnGUI()
    {
        GUILayout.Label("Base Settings", EditorStyles.boldLabel);

        if (GUILayout.Button("Select Input Path"))
        {
            string defaultPath = string.IsNullOrEmpty(inputPath) ? "" : inputPath;
            inputPath = EditorUtility.OpenFolderPanel("Select Input Folder", defaultPath, "");
        }

        EditorGUILayout.LabelField("Input Path: " + inputPath);

        if (GUILayout.Button("Select Output Path"))
        {
            string defaultPath = string.IsNullOrEmpty(outputPath) ? "" : outputPath;
            outputPath = EditorUtility.OpenFolderPanel("Select Output Folder", defaultPath, "");
        }

        EditorGUILayout.LabelField("Output Path: " + outputPath);

        lineCount = EditorGUILayout.IntField("Line Count: ", lineCount);

        if (GUILayout.Button("Convert"))
        {
            EditorPrefs.SetString("inputPathXTC", inputPath);
            EditorPrefs.SetString("outputPathXTC", outputPath);
            EditorPrefs.SetInt("lineCountXTC", lineCount);

            string projectPath = Application.dataPath.Replace('/', '\\');

            ConvertFiles(inputPath, outputPath);
        }
    }

    private void OnEnable()
    {
        inputPath = EditorPrefs.GetString("inputPathXTC", "");
        outputPath = EditorPrefs.GetString("outputPathXTC", "");
        lineCount = EditorPrefs.GetInt("lineCountXTC", 0);
    }

    private void ConvertFiles(string inputPath, string outputPath)
    {
        string[] fileEntries = Directory.GetFiles(inputPath, "*.xlsx", SearchOption.AllDirectories);

        Debug.Log("fileEntries.Count   :" + fileEntries.Length);

        foreach (string fileName in fileEntries)
        {
            string subPath = Path.GetRelativePath(inputPath, fileName);

            string outputSubFolder = Path.GetDirectoryName(subPath);

            string outputFolderFullPath = Path.Combine(outputPath, outputSubFolder);
            Directory.CreateDirectory(outputFolderFullPath);

            string outputFileNameWithoutExtension = Path.GetFileNameWithoutExtension(subPath);

            string outputCsvFile = Path.Combine(outputFolderFullPath, outputFileNameWithoutExtension + ".csv");

            if (File.Exists(outputCsvFile))
            {
                File.Delete(outputCsvFile);
            }
            if (!csvFileOuts.ContainsKey(fileName))
            {
                csvFileOuts.Add(fileName, outputCsvFile);
            }
        }

        //string shieldNameFilePath = Path.Combine(Application.dataPath, "Resources", "shieldName.xlsx");
        //shieldNames = GetShieldNames(shieldNameFilePath);

        EachXlsxToCsv();
    }

    private void EachXlsxToCsv()
    {
        foreach (KeyValuePair<string, string> item in csvFileOuts)
        {
            ConvertXlsxToCsvAc(item.Key, item.Value, lineCount);
        }
        EditorUtility.DisplayDialog("Conversion Complete", "XLSX to CSV conversion is complete.", "OK");
    }
    /// <summary>
    /// 兼容WPS以及各种版本的Excel表格
    /// </summary>
    private void ConvertXlsxToCsvAc(string inputFileName, string outputFileName, int lineCount = 3)
    {
        if (string.IsNullOrEmpty(inputFileName) || string.IsNullOrEmpty(outputFileName))
        {
            Debug.LogError("File paths cannot be null or empty.");
            return;
        }

        // 排除以 "~$" 开头的临时文件
        if (Path.GetFileName(inputFileName).StartsWith("~$"))
        {
            Debug.LogWarning("Skipping temporary file: " + inputFileName);
            return;
        }

        string tempFilePath = Path.GetTempFileName();
        File.Copy(inputFileName, tempFilePath, true);       //即便表格在打开的状态下，也可进行转表

        using (FileStream file = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = null;
            try
            {
                workbook = new XSSFWorkbook(file);
            }
            catch (System.Exception)
            {
                workbook = new HSSFWorkbook(file);  //Excel2003 的表格，采用此格式读取
            }

            ISheet sheet = workbook.GetSheetAt(0);

            int maxCellNum = 0;
            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null && row.LastCellNum > maxCellNum)
                {
                    maxCellNum = row.LastCellNum;
                }
            }

            int csvFileIndex = 1;
            StringBuilder csvContent = new StringBuilder();
            List<string> headerRows = new List<string>();

            for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);

                if (row == null)
                {
                    csvContent.AppendLine();
                    continue;
                }

                List<string> cells = new List<string>();

                for (int j = 0; j < maxCellNum; j++)
                {
                    ICell cell = row.GetCell(j);

                    string cellValue = GetCellValue(cell);
                    cells.Add(cellValue);
                }

                if (i < 3)
                {
                    headerRows.Add(string.Join(",", cells));
                    continue;
                }

                csvContent.AppendLine(string.Join(",", cells));

                int mul = lineCount > 3 ? (i - 2) % lineCount : 1;

                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(outputFileName);

                //if (!shieldNames.Contains(fileNameWithoutExtension))
                //{
                  
                //}

                if (mul == 0 && i > 2)
                {
                    string directoryPath = Path.GetDirectoryName(outputFileName);

                    string baseOutputFileName = Path.Combine(directoryPath, fileNameWithoutExtension);
                    File.WriteAllText(baseOutputFileName + "_" + csvFileIndex + ".csv", string.Join("\r\n", headerRows) + "\r\n" + csvContent.ToString());
                    csvContent.Clear();
                    csvFileIndex++;
                }
            }

            if (csvContent.Length > 0)
            {
                if (csvFileIndex == 1)
                {
                    File.WriteAllText(outputFileName, string.Join("\r\n", headerRows) + "\r\n" + csvContent.ToString(), new UTF8Encoding(true));
                }
                else
                {
                    string directoryPath = Path.GetDirectoryName(outputFileName);
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(outputFileName);
                    string baseOutputFileName = Path.Combine(directoryPath, fileNameWithoutExtension);
                    File.WriteAllText(baseOutputFileName + "_" + csvFileIndex + ".csv", string.Join("\r\n", headerRows) + "\r\n" + csvContent.ToString(), new UTF8Encoding(true));
                }

            }
        }
        File.Delete(tempFilePath);
    }

    private string GetCellValue(ICell cell)
    {
        if (cell == null)
            return "";

        switch (cell.CellType)
        {
            case CellType.Blank:
                return "";
            case CellType.Boolean:
                return cell.BooleanCellValue.ToString();
            case CellType.Error:
                return "";
            case CellType.Formula:
                switch (cell.CachedFormulaResultType)
                {
                    case CellType.Blank:
                        return "";
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return "";
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    default:
                        return "";
                }
            case CellType.Numeric:
                return cell.NumericCellValue.ToString();
            case CellType.String:
                return cell.StringCellValue;
            default:
                return "";
        }
    }

   
}