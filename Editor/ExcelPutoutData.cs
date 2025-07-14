using System.Collections;
using System.Collections.Generic;
using UnityEngine;

namespace ExcelToJson.Editor
{
    public class ExcelPutoutData
    {
        /// <summary>
        /// json���ݣ���һ��keyΪid���ڶ���keyΪ��������
        /// </summary>
        public Dictionary<string, Dictionary<string, object>> JsonData;
        /// <summary>
        /// �����б�keyΪ�������ƣ�valueΪ��������
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