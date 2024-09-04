using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using UnityEditor;
using UnityEngine;
using Debug = UnityEngine.Debug;

public class XlsxToJson : EditorWindow
{
    string inputPath = "";
    string outputPath = "";

    [MenuItem("XlsxTools/Õ®”√Xlsx to Json")]
    public static void ShowWindow()
    {
        EditorWindow.GetWindow(typeof(XlsxToJson));
    }

    Dictionary<string, string> csvFileOuts = new Dictionary<string, string>();
    List<string> jsonFlors = new List<string>();
    string batFilePath = "";
    string exePath = "";

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

        if (GUILayout.Button("Convert"))
        {
            EditorPrefs.SetString("inputPathXTJ", inputPath);
            EditorPrefs.SetString("outputPathXTJ", outputPath);

            string projectPath = Application.dataPath.Replace('/', '\\');
            batFilePath = Path.Combine(projectPath, "..", "Tools", "exceltojson5.bat");
            exePath = Path.Combine(projectPath, "..", "Tools", "excel2json.exe");

            ConvertFiles(inputPath, outputPath);

            //RunBatFile(batFilePath, inputPath, outputPath);
        }
    }

    private void ConvertFiles(string inputPath, string outputPath)
    {
        string[] fileEntries = Directory.GetFiles(inputPath, "*.xlsx", SearchOption.AllDirectories);

        //string curOutFolder = "";
        foreach (string fileName in fileEntries)
        {
            string subPath = Path.GetRelativePath(inputPath, fileName);

            string outputSubFolder = Path.GetDirectoryName(subPath);
            string outputFolderFullPath = Path.Combine(outputPath, outputSubFolder);
            Directory.CreateDirectory(outputFolderFullPath);
            string outputFileNameWithoutExtension = Path.GetFileNameWithoutExtension(subPath);

            {
                string outputJsonFile = Path.Combine(outputFolderFullPath, outputFileNameWithoutExtension + ".text");

                if (File.Exists(outputJsonFile))
                {
                    File.Delete(outputJsonFile);
                }

                string excelFilePath = Path.Combine(inputPath, outputSubFolder);

                Debug.Log("json excelFilePath   :" + excelFilePath);

                //string Arguments = "\"" + excelFilePath + "\" \"" + outputFolderFullPath + "\"";
                string Arguments = $"\"{excelFilePath}\" \"{outputFolderFullPath}\" \"{exePath}\"";

                if (!jsonFlors.Contains(Arguments))
                {
                    jsonFlors.Add(Arguments);
                }
            }
        }
        RunBatFile();
    }

    private void OnEnable()
    {
        inputPath = EditorPrefs.GetString("inputPathXTJ", "");
        outputPath = EditorPrefs.GetString("outputPathXTJ", "");
    }

    public void RunBatFile()//(string batFilePath, string inAndOutPath)//string excelFilePath, string outputFolder)
    {
        if (jsonFlors.Count <= 0)
        {
            return;
        }

        Process process = new Process();
        process.StartInfo.FileName = batFilePath;
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = true; // Redirect the standard error stream

        string inAndOutPath = jsonFlors[jsonFlors.Count - 1];

        process.StartInfo.Arguments = inAndOutPath;//"\"" + excelFilePath + "\" \"" + outputFolder + "\""; // pass the excel file path and output folder as arguments
        process.Start();

        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();

        process.WaitForExit();

        if (!string.IsNullOrEmpty(output))
        {
            UnityEngine.Debug.Log(output);
        }

        if (!string.IsNullOrEmpty(error))
        {
            UnityEngine.Debug.LogError(error);
        }

        jsonFlors.RemoveAt(jsonFlors.Count - 1);

        RunBatFile();
    }
}
