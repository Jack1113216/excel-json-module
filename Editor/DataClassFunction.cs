using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityExtension;

namespace ExcelToJson.Editor
{
    public class DataClassFunction
    {
        static string classStr = "using JsonConfig;\nusing UnityExtension;\n\nnamespace Config\n{\n    public class {className} : IConfig\n    {{variables}}\n}";
        static string variableStr = "\n        public {type} {name}";
        public static string CreateClassStr(string className, Dictionary<string, string> variables)
        {
            var str = "";
            foreach (var variable_kv in variables)
            {
                str += getVariables(variable_kv);
            }
            if (str != "")
                str += "\n    ";
            return classStr.Replace("{className}", className).Replace("{variables}", str);
        }
        static string getVariables(KeyValuePair<string, string> variable_kv)
        {
            var str = variableStr.Replace("{name}", variable_kv.Key).Replace("{type}", variable_kv.Value);
            str += variable_kv.Key == "ID" ? " { get; set; }" : ";";
            return str;
        }
    }
}