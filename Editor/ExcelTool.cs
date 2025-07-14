using EditorFileManager;
using System.Collections;
using System.IO;
using Unity.EditorCoroutines.Editor;
using UnityEditor;
using UnityEngine;
using UnityExtension;
namespace ExcelToJson.Editor
{
    public class ExcelTool
    {
        /// <summary>
        /// excel文件夹路径
        /// </summary>
        static string excel_root = Application.dataPath.Replace("/Assets", "/Excel");
        /// <summary>
        /// 输出的json文件路径
        /// </summary>
        static string json_root = Application.dataPath + "/Resources/Configs/";
        /// <summary>
        /// data类型脚本存放路径
        /// </summary>
        static string script_root = Application.dataPath + "/Script/Common/Data/";

        /// <summary>
        /// 菜单栏中，excel转json工具按钮
        /// </summary>
        [MenuItem("Tools/Excel to Json", false, 2000)]
        static void ExcelToJson()
        {
            EditorCoroutineUtility.StartCoroutineOwnerless(CreateFileIE());
        }
        /// <summary>
        /// 创建文件
        /// </summary>
        /// <returns></returns>
        static IEnumerator CreateFileIE()
        {

            var file_list = Directory.GetFiles(excel_root);
            foreach (var file in file_list)
            {
                var fileName = Path.GetFileNameWithoutExtension(file);
                using (var excel = ExcelManager.LoadExcel(file))
                {
                    if (excel != null)
                    {
                        foreach (var sheet in excel.Workbook.Worksheets)
                        {
                            var sheetName = sheet.Name.ToCamel();
                            yield return ExcelManager.GetExcelPutout(sheet, (data) =>
                            {
                                SaveJson(data, sheetName);
                                SaveClass(data, sheetName);
                            });
                        }
                    }
                }
                Debug.Log($"{fileName} excel create succes");
            }
            Debug.Log("all Excel Created");
            AssetDatabase.Refresh();
            yield break;
        }
        /// <summary>
        /// 创建json文件
        /// </summary>
        /// <param name="data">excel数据</param>
        /// <param name="fileName">文件名</param>
        static void SaveJson(ExcelPutoutData data, string fileName)
        {
            var json = data.JsonData.ToJson();
            string path = json_root + fileName + ".json";
            EditorFileMgr.CreateTextFile(path, json, true);
        }
        /// <summary>
        /// 创建脚本文件
        /// </summary>
        /// <param name="data">excel数据</param>
        /// <param name="fileName">文件名</param>
        static void SaveClass(ExcelPutoutData data, string fileName)
        {
            if (!fileName.EndsWith("_Config"))
                fileName = fileName + "_Config";
            var script = DataClassFunction.CreateClassStr(fileName, data.variableList);
            EditorFileMgr.CreateTextFile(script_root + fileName + ".cs", script, true);
        }
    }
}