using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml;

namespace Install_From_MSI
{
     class GeneralConfig {

        private readonly string fileName = "Config.xml";
        public  string PathToDirLog { get; private set; }
        public string PathToDirXML { get; private set; }
        public  Dictionary<string,string> PtrnNames { get; private set; }

        private GeneralConfig() {
            PathToDirLog = String.Empty;
            PtrnNames = new Dictionary<string, string>();
            GetConfig(fileName);
        }
        private static GeneralConfig generalconfig = new GeneralConfig();
        public static GeneralConfig GetInstance() {
            return generalconfig;
        }
        
        private  void GetConfig(string filename) {
            string Pathfile = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), filename);
            PathToDirXML = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "XML");

            Dictionary<string, Action<XmlNode>> dic = new Dictionary<string, Action<XmlNode>>
            {
                ["PathToDirLog"] = x => { this.PathToDirLog = x.InnerText.ToString(); },
                ["PathToDirXML"] = x => { this.PathToDirXML = CheckDirectory(x.InnerText.ToString())? x.InnerText.ToString() : this.PathToDirXML; },
                ["Patterns"] = x =>
                {
                    foreach (XmlNode patternnode in x)
                    {
                        this.PtrnNames.Add(patternnode.Attributes["appname"].Value.ToString(), patternnode.InnerText.ToString());
                    }
                }
            };
            
            try
            {
                using (Stream stream = new FileStream(Pathfile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(stream);
                    XmlElement XRoot = xDoc.DocumentElement;
                    foreach (XmlNode xnode in XRoot)
                    {
                        if (dic.ContainsKey(xnode.Name.ToString()))                        
                            dic[xnode.Name.ToString()].Invoke(xnode);                        
                    }
                }
            }
            catch (Exception e)
            {
                //check here why it failed and ask user to retry if the file is in use.
                Console.WriteLine(e.Message);
            }

        }

        private bool CheckDirectory(string pathToDir) {
            if (String.IsNullOrEmpty(pathToDir)) {
                Console.WriteLine($"Warning: PathToDirXML in Config.xml - not set.");
                return false;
            }
            if (Directory.Exists(pathToDir)) {
                return true;
            }
            Console.WriteLine($"Error: path to directory 'XML' {pathToDir}  from Config.xml NOT exists!!!");
            return false;        
        }
    }

}
