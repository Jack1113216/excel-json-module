using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using UnityEngine;
using UnityEngine.Events;
using UnityExtension;

namespace ExcelToJson.Editor
{
    public static class ExcelManager
    {
        /// <summary>
        /// 变量名所在行数
        /// </summary>
        public static int keyRow = 1;
        /// <summary>
        /// 变量类型所在行数
        /// </summary>
        public static int typeRow = 2;
        /// <summary>
        /// 数据起始行数
        /// </summary>
        public static int startRow = 3;
        /// <summary>
        /// 从路径中读取excel文件
        /// </summary>
        /// <param name="path">excel路径</param>
        /// <returns></returns>
        public static ExcelPackage LoadExcel(string path)
        {
            if (File.Exists(path) && path.EndsWith(".xlsx"))
            {
                var file = new FileInfo(path);
                try
                {
                    var excel = new ExcelPackage(file);
                    return excel;
                }
                catch
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// 根据excel中的表格数据，变成
        /// </summary>
        /// <param name="sheet">sheet表单</param>
        /// <param name="callback">获取到输出数据后的回调</param>
        /// <returns></returns>
        public static IEnumerator GetExcelPutout(ExcelWorksheet sheet, UnityAction<ExcelPutoutData> callback)
        {
            //var sheetName = sheet.Name;

            if (!RooterName(sheet.Name))
            {
                Debug.Log($"{sheet.Name} is not rooter");
                yield break;
            }
            List<string> key_list = new();
            List<string> type_list = new();
            int i = 1;
            while (sheet.Cells[keyRow, i].Value != null)
            {
                var key = sheet.Cells[keyRow, i].Value.ToString();
                if (string.IsNullOrEmpty(key) || !RooterName(key))
                {
                    Debug.Log(key + " is error");
                    break;
                }
                key = key.ToCamel();
                if (key == "Id" || key == "ID")
                    key = "ID";
                else
                    key = char.ToLower(key[0]) + key[1..];
                var type = sheet.Cells[typeRow, i].Value.ToString();
                key_list.Add(key);
                type_list.Add(GetVariableType(type));
                i++;
            }
            ExcelPutoutData putoutData = new(key_list, type_list);
            //从起始行开始，逐行读取数据并赋值
            int n = startRow;
            while (sheet.Cells[n, 1].Value != null)
            {
                var id = sheet.Cells[n, 1].Value.ToString();
                putoutData.JsonData.Add(id, new());
                for (int m = 1; m <= key_list.Count; m++)
                {
                    var key = key_list[m - 1];
                    var str = sheet.Cells[n, m].Value.ToString() ?? "";
                    var value = GetValue(str, type_list[m - 1]);
                    putoutData.JsonData[id].Add(key, value);
                }
                n++;
            }
            callback?.Invoke(putoutData);
            yield break;
        }
        /// <summary>
        /// 判断名称是否符合c#变量命名规范
        /// </summary>
        /// <param name="str">源字符串</param>
        /// <returns></returns>
        static bool RooterName(string str)
        {
            var isNumberLetter = Regex.IsMatch(str, @"^[a-zA-Z0-9_ ]+$");
            var startNumber = Regex.IsMatch(str, @"^[0-9]+");
            return isNumberLetter && !startNumber;
        }
        /// <summary>
        /// 转换数据类型
        /// </summary>
        /// <param name="str">源数据类型名</param>
        /// <returns></returns>
        static string GetVariableType(string str)
        {
            str = str.ToLower();
            var type = str.Replace("array", "[]");
            if (Regex.IsMatch(str, @"^v\d+.*"))
            {
                type = type.Replace("v", "DataVect");
            }
            return type;
        }
        /// <summary>
        /// 文字转为数据
        /// </summary>
        /// <param name="str">源字符串</param>
        /// <param name="type">数据类型</param>
        /// <returns></returns>
        static object GetValue(string str, string type)
        {
            object value = null;
            if (type.EndsWith("[]"))//转为数组
            {
                Debug.Log(type + "???" + str);
                value = GetValueArray(str, type[..^2]);
            }
            else if (type == "int")//转为int
            {
                value = int.Parse(str);
            }
            else if (type == "float")//转为float
            {
                value = float.Parse(str);
            }
            else if (type == "string")//转为string
            {
                value = str;
            }
            else if (type == "DataVect2")//转为datavect2
            {
                Debug.Log(type + ">>>>>>>" + str);
                value = DataVect2.Parse(str);
            }
            else if (type == "DataVect3")//转为datavect3
            {
                value = DataVect3.Parse(str);
            }
            else if (type == "DataVect4")//转为datavect4
            {
                value = DataVect4.Parse(str);
            }
            return value;
        }
        /// <summary>
        /// 文字转数组
        /// </summary>
        /// <param name="str">源字符串</param>
        /// <param name="type">数组类型</param>
        /// <returns></returns>
        static List<object> GetValueArray(string str, string type)
        {
            var itemArray = str.Split(new char[] { '|', ';', '(', ')', '；', '(', ')' }, System.StringSplitOptions.RemoveEmptyEntries);
            List<object> objList = new();
            foreach (var item in itemArray)
            {
                var obj = GetValue(item, type);
                if (obj != null)
                    objList.Add(obj);
            }
            return objList;
        }
    }
}