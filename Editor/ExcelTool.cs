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
        /// excel�ļ���·��
        /// </summary>
        static string excel_root = Application.dataPath.Replace("/Assets", "/Excel");
        /// <summary>
        /// �����json�ļ�·��
        /// </summary>
        static string json_root = Application.dataPath + "/Resources/Configs/";
        /// <summary>
        /// data���ͽű����·��
        /// </summary>
        static string script_root = Application.dataPath + "/Script/Common/Data/";

        /// <summary>
        /// �˵����У�excelתjson���߰�ť
        /// </summary>
        [MenuItem("Tools/Excel to Json", false, 2000)]
        static void ExcelToJson()
        {
            EditorCoroutineUtility.StartCoroutineOwnerless(CreateFileIE());
        }
        /// <summary>
        /// �����ļ�
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
        /// ����json�ļ�
        /// </summary>
        /// <param name="data">excel����</param>
        /// <param name="fileName">�ļ���</param>
        static void SaveJson(ExcelPutoutData data, string fileName)
        {
            var json = data.JsonData.ToJson();
            string path = json_root + fileName + ".json";
            EditorFileMgr.CreateTextFile(path, json, true);
        }
        /// <summary>
        /// �����ű��ļ�
        /// </summary>
        /// <param name="data">excel����</param>
        /// <param name="fileName">�ļ���</param>
        static void SaveClass(ExcelPutoutData data, string fileName)
        {
            if (!fileName.EndsWith("_Config"))
                fileName = fileName + "_Config";
            var script = DataClassFunction.CreateClassStr(fileName, data.variableList);
            EditorFileMgr.CreateTextFile(script_root + fileName + ".cs", script, true);
        }
    }
}