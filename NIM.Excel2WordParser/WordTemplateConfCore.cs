using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace NIM.CertificationGenerator
{
    public class WordTemplateConfCore 
    {
        public string Code
        {
            get;
            private set;
        }
        public string FileFullName
        {
            get;
            private set;
        }


        private  Dictionary<string, string> mProductCertificationPaths = new Dictionary<string, string>();



        private  Dictionary<string, string> GetPaths(string confFilePath)
        {
            var dictornies = new Dictionary<string, string>();

            if (!System.IO.File.Exists(confFilePath))
                throw new Exception($"找不到证书模板配置文件.({confFilePath})");
            var str = File.ReadAllText(confFilePath);

            var wordConfBasePath = System.IO.Path.GetDirectoryName(confFilePath);
            var xml = XDocument.Parse(str);
            var root = xml.Root;
            //  < item productname = "辐射温度计" filename = "rr-RGrr2017.docx" >
            var items = root.Elements().ToList();
            items.ForEach(t =>
            {
                var code = t.Attribute("code").Value;
                var filename = t.Attribute("filename").Value;
                if (dictornies.ContainsKey(code))
                    throw new Exception($"重复的Code在配置文件中({confFilePath},{code}).");
                filename = Path.Combine(wordConfBasePath, filename);
                if (!System.IO.File.Exists(filename))
                    throw new Exception($"配置文件指定的word模板文件不存在({confFilePath},{filename}).");
                dictornies.Add(code, filename);

            });

            return dictornies;
        }

        private  string getConfFilePath(string excelFilePath)
        {
            var wordConfPath = IFilePathManagerProvider.PathProvider.WordTemplateConfPath; 

            var confFilePath = Path.Combine(wordConfPath, "conf.xml");
            return confFilePath;
        }
        public  string GetCertificateTemplateConfFileFullName(string productName, string excelFileFullName)
        {

            if (mProductCertificationPaths.Keys.Count == 0)
            {
                lock (mProductCertificationPaths)
                {
                    var confFilePath = getConfFilePath(excelFileFullName);
                    mProductCertificationPaths = GetPaths(confFilePath);
                }
            }
            if (!mProductCertificationPaths.ContainsKey(productName))
                throw new Exception($"在证书模板配置文件中，根据产品名称找不到指定的模板({getConfFilePath(excelFileFullName)},{productName}).");

            return mProductCertificationPaths[productName];

        }
    }
}
