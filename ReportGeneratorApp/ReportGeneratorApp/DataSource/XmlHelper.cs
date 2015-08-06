using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace ReportGeneratorApp.DataSource
{
    class XmlHelper
    {
        public static string GetConfigFilePath(string shortName)
        {
            return string.Format(@"{0}Config\{1}Config.xml", AppDomain.CurrentDomain.BaseDirectory, shortName);
        }
        public static string GetSql(string shortName, string key)
        {
            XElement doc = XElement.Load(GetConfigFilePath(shortName));
            var sqls = doc.Descendants("sqls").Elements();
            return sqls.Where(w => w.Attribute("key").Value == key).FirstOrDefault().Value.Trim();
        }

        public static Dictionary<string, object> GetProcess(string shortName)
        {
            string configFile = GetConfigFilePath(shortName);
            if (!File.Exists(configFile)) return null;
            XElement doc = XElement.Load(configFile);
            var processes = doc.Descendants("process").Elements();
            if (processes.Count() == 0) return null;
            Dictionary<string, object> retval = new Dictionary<string, object>();
            int i = 0;
            foreach (var process in processes)
            {
                if (process.Attribute("type") == null) continue;
                Dictionary<string, object> paramList = new Dictionary<string, object>();
                foreach (var attribute in process.Attributes())
                {
                    paramList.Add(attribute.Name.LocalName, attribute.Value);
                }
                if (process.HasElements)
                {
                    //int j = 0;
                    //Dictionary<string, object> subParams = new Dictionary<string, object>();
                    //foreach (var element in process.Elements())
                    //{
                    //    Dictionary<string, object> subParamsAttrList = new Dictionary<string, object>();
                    //    foreach (var attribute in element.Attributes())
                    //    {
                    //        subParamsAttrList.Add(attribute.Name.LocalName, attribute.Value);
                    //    }
                    //    subParams.Add(j.ToString(), subParamsAttrList);
                    //    j++;
                    //}
                    paramList.Add("SubParams", process.Elements());
                }
                retval.Add(i.ToString(), paramList);
                i++;
            }
            return retval;
        }

        public static Dictionary<string, object> GetFillParams(string shortName, string key)
        {
            string configFile = GetConfigFilePath(shortName);
            if (!File.Exists(configFile)) return null;
            XElement doc = XElement.Load(configFile);
            var fillParams = doc.Descendants("fillParams").Elements();
            if (fillParams.Count() == 0) return null;
            Dictionary<string, object> retval = new Dictionary<string, object>();
            foreach (var fillParam in fillParams)
            {
                if (fillParam.Attribute("key") == null || fillParam.Attribute("key").Value.ToString() != key) continue;
                foreach (var attribute in fillParam.Attributes())
                {
                    if (attribute.Name.LocalName == "key") continue;
                    retval.Add(attribute.Name.LocalName, attribute.Value);
                }
            }
            return retval;
        }

        public static Dictionary<string, object> GetProcParams(string shortName, string key)
        {
            string configFile = GetConfigFilePath(shortName);
            if (!File.Exists(configFile)) return null;
            XElement doc = XElement.Load(configFile);
            var procParams = doc.Descendants("procParams").Elements();
            if (procParams.Count() == 0) return null;
            Dictionary<string, object> retval = new Dictionary<string, object>();
            foreach (var procParam in procParams)
            {
                if (procParam.Attribute("key") == null || procParam.Attribute("key").Value.ToString() != key) continue;
                foreach (var attribute in procParam.Attributes())
                {
                    if (attribute.Name.LocalName == "key") continue;
                    retval.Add(attribute.Name.LocalName, attribute.Value);
                }
                if (procParam.HasElements)
                {
                    retval.Add("paramsList", procParam.Elements());
                }
            }
            return retval;
        }
    }
}
