using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Singleton;
using System.Collections.Generic;
using System.Linq;
using UnityEngine;
using UnityExtension;

namespace JsonConfig
{
    /// <summary>
    /// 配置管理器
    /// </summary>
    public class ConfigMgr : SingletonMgr<ConfigMgr>
    {
        /// <summary>
        /// 已加载的配置
        /// </summary>
        Dictionary<string, Dictionary<string, IConfig>> configDics;
        private ConfigMgr()
        {
            configDics = new();
        }
        /// <summary>
        /// 获取配置
        /// </summary>
        /// <typeparam name="T">配置类型</typeparam>
        /// <returns></returns>
        public Dictionary<string, T> GetConfig<T>() where T : IConfig
        {
            var configName = GetConfigName<T>();
            Dictionary<string, T> dic;
            if (configDics.ContainsKey(configName))
            {
                dic = configDics[configName].ToDictionary(pair => pair.Key, pair => (T)pair.Value);
            }
            else
            {
                var path = $"Configs/{configName}";
                var text = Resources.Load<TextAsset>(path).text;
                dic = text.ToObject<Dictionary<string, T>>();
                var configList = dic.ToDictionary(pair => pair.Key, pair => (IConfig)pair.Value);
                configDics.Add(configName, configList);
            }
            return dic;
        }
        /// <summary>
        /// 移除配置缓存
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public void RemoveConfig<T>()
        {
            var configName = GetConfigName<T>();
            if (configDics.ContainsKey(configName))
            {
                configDics.Remove(configName);
            }
        }
        /// <summary>
        /// 获取配置名
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        string GetConfigName<T>()
        {
            var configName = typeof(T).Name;
            Debug.Log(configName);
            return configName[..^7];
        }
    }
}