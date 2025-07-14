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
        /// ��������������
        /// </summary>
        public static int keyRow = 1;
        /// <summary>
        /// ����������������
        /// </summary>
        public static int typeRow = 2;
        /// <summary>
        /// ������ʼ����
        /// </summary>
        public static int startRow = 3;
        /// <summary>
        /// ��·���ж�ȡexcel�ļ�
        /// </summary>
        /// <param name="path">excel·��</param>
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
        /// ����excel�еı�����ݣ����
        /// </summary>
        /// <param name="sheet">sheet��</param>
        /// <param name="callback">��ȡ��������ݺ�Ļص�</param>
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
            //����ʼ�п�ʼ�����ж�ȡ���ݲ���ֵ
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
        /// �ж������Ƿ����c#���������淶
        /// </summary>
        /// <param name="str">Դ�ַ���</param>
        /// <returns></returns>
        static bool RooterName(string str)
        {
            var isNumberLetter = Regex.IsMatch(str, @"^[a-zA-Z0-9_ ]+$");
            var startNumber = Regex.IsMatch(str, @"^[0-9]+");
            return isNumberLetter && !startNumber;
        }
        /// <summary>
        /// ת����������
        /// </summary>
        /// <param name="str">Դ����������</param>
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
        /// ����תΪ����
        /// </summary>
        /// <param name="str">Դ�ַ���</param>
        /// <param name="type">��������</param>
        /// <returns></returns>
        static object GetValue(string str, string type)
        {
            object value = null;
            if (type.EndsWith("[]"))//תΪ����
            {
                Debug.Log(type + "???" + str);
                value = GetValueArray(str, type[..^2]);
            }
            else if (type == "int")//תΪint
            {
                value = int.Parse(str);
            }
            else if (type == "float")//תΪfloat
            {
                value = float.Parse(str);
            }
            else if (type == "string")//תΪstring
            {
                value = str;
            }
            else if (type == "DataVect2")//תΪdatavect2
            {
                Debug.Log(type + ">>>>>>>" + str);
                value = DataVect2.Parse(str);
            }
            else if (type == "DataVect3")//תΪdatavect3
            {
                value = DataVect3.Parse(str);
            }
            else if (type == "DataVect4")//תΪdatavect4
            {
                value = DataVect4.Parse(str);
            }
            return value;
        }
        /// <summary>
        /// ����ת����
        /// </summary>
        /// <param name="str">Դ�ַ���</param>
        /// <param name="type">��������</param>
        /// <returns></returns>
        static List<object> GetValueArray(string str, string type)
        {
            var itemArray = str.Split(new char[] { '|', ';', '(', ')', '��', '(', ')' }, System.StringSplitOptions.RemoveEmptyEntries);
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