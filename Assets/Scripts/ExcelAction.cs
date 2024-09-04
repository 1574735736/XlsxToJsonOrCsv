using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
using System;

#if UNITY_EDITOR
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
#endif

using Dialogue;

public class ExcelAction : ScriptableObject {
    public string saveScriptableObjectPath; //对话数据文件夹存储路径
    public string excelPath;                //excel表格文件路径

    public void WriteExcel()
    {
        //空值判断
        if (string.IsNullOrEmpty(excelPath))
            return;
        if (!excelPath.Contains(".xlsx"))
        {
            Debug.Log("路径不是excel文件");
            return;
        }
        //打开文件流
        FileStream fs = File.Exists(excelPath) ? File.Open(excelPath, FileMode.Open) : File.Create(excelPath);
        //新建excel和sheet
        var wk = new HSSFWorkbook();
        var sheet = wk.CreateSheet();

        int i = 1; //省略掉第0行

        //获取指定目录下的所有文件
        var dialogueDatas = GetFiles<DialogueData>(saveScriptableObjectPath);
        Debug.Log("加载" + dialogueDatas.Count + "个对话");
        
        //遍历所有对话文件
        foreach (var dialogue in dialogueDatas)
        {
            var row = sheet.CreateRow(i);
            var cell = row.CreateCell(0);
            //讲ScriptableObject的name写在第一列
            cell.SetCellValue(dialogue.name);
            //遍历句子内容
            for (int j = 0; j < dialogue.data.Count; j++)  
            {
                if (j != 0)
                    row = sheet.CreateRow(i);
                //将名字写在第二列
                row.CreateCell(1).SetCellValue(dialogue.data[j].character.ToString());
                //将对话内容写在第三列
                row.CreateCell(2).SetCellValue(dialogue.data[j].content);
                i++;
            }
        }


        wk.Write(fs);
        fs.Close();
        fs.Dispose();
        Debug.Log("写入成功");
    }
    
    public void ReadExcel()
    {
        //空值判断
        if (string.IsNullOrEmpty(excelPath))
            return;
        if (!excelPath.Contains(".xlsx"))
        {
            Debug.Log("路径不是excel文件");
            return;
        }
        //打开文件流
        FileStream fs = File.Open(excelPath, FileMode.Open);
        //打开excel和sheet
        var wk = new HSSFWorkbook(fs);
        var sheet = wk.GetSheetAt(0);
        //这是ScriptableObject的实例
        DialogueData so = null;
        //开始遍历sheet每一行的数据，注意这里i=1跳过了第一行
        for (int i = 1; i < sheet.LastRowNum; i++)
        {
            Debug.Log("读取EXCEL 行数:" + i);
            var row = sheet.GetRow(i);
            //如果当前行第一列元素不为空，证明需要保存当前so，然后创建新的so
            if (row.GetCell(0) != null && !string.IsNullOrEmpty(row.GetCell(0).ToString()))
            {
                //只有第一次so为null
                if (so) //保存路径
                {
                    //是否原来已经创建了ScriptableObject资源
                    if (File.Exists(saveScriptableObjectPath + "/" + so.name + ".asset"))
                    {
                        var loadSo = AssetDatabase.LoadAssetAtPath<DialogueData>(saveScriptableObjectPath + "/" + so.name + ".asset");
                        loadSo.data = so.data;
                        EditorUtility.SetDirty(loadSo); //标记为已修改
                    }
                    else
                        AssetDatabase.CreateAsset(so, saveScriptableObjectPath + "/" + so.name + ".asset");
                }
                //创建实例，开始新的对话数据记录
                so = ScriptableObject.CreateInstance<DialogueData>();
                so.name = row.GetCell(0).ToString();
            }
            //加载excel第二第三列的数据
            if (so != null)
            {
                var se = new DialogueData.Sentence();
                se.character = row.GetCell(1).ToString();
                se.content = row.GetCell(2).ToString();
                so.data.Add(se);
            }
        }
        //退出循环时记得还有一个so的数据
        if (so)
        {
            if (File.Exists(saveScriptableObjectPath + "/" + so.name + ".asset"))
            {
                File.Delete(saveScriptableObjectPath + "/" + so.name + ".meta");
                File.Delete(saveScriptableObjectPath + "/" + so.name + ".asset");
            }
            AssetDatabase.CreateAsset(so, saveScriptableObjectPath + "/" + so.name + ".asset");
        }

        AssetDatabase.SaveAssets();
        AssetDatabase.Refresh();
    }
    
    //获取指定路径下面的所有资源文件  
    public static List<T> GetFiles<T>(string dir) where T : UnityEngine.Object
    {
        string path = string.Format(dir);
        var list = new List<T>();
        if (Directory.Exists(path))
        {
            DirectoryInfo direction = new DirectoryInfo(path);
            FileInfo[] files = direction.GetFiles("*");

            for (int i = 0; i < files.Length; i++)
            {
                //忽略关联文件
                if (files[i].Name.EndsWith(".meta"))
                {
                    continue;
                }
#if UNITY_EDITOR      
                var so = AssetDatabase.LoadAssetAtPath<T>(dir + "/" + files[i].Name);
                if (so != null)
                {
                    Debug.Log("加载资源" + files[i].Name);
                    list.Add(so as T);
                }
#endif
            }
        }

        return list;
    }
}
