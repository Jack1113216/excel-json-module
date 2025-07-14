using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace ExcelToJson.Editor
{
    public class ExcelPutoutData
    {
        /// <summary>
        /// json数据，第一个key为id，第二个key为变量名称
        /// </summary>
        public Dictionary<string, Dictionary<string, object>> JsonData;
        /// <summary>
        /// 变量列表，key为变量名称，value为变量类型
        /// </summary>
        public Dictionary<string, string> variableList;
        public ExcelPutoutData()
        {
            JsonData = new();
            variableList = new();
        }
        public ExcelPutoutData(List<string> key, List<string> type)
        {
            JsonData = new();
            variableList = new();
            for (int i = 0; i < key.Count; i++)
            {
                variableList.Add(key[i], type[i]);
            }
        }
        public void SetClass(List<string> key, List<string> type)
        {
            for (int i = 0; i < key.Count; i++)
            {
                variableList.Add(key[i], type[i]);
            }
        }
    }
}