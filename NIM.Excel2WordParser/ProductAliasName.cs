using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator
{
    public class ProductAliasName
    {
        public string Name
        {
            get; set;
        }
        public List<string> AliasName
        {
            get; set;
        }

        private static List<ProductAliasName> _names;
        static ProductAliasName()
        {
            _names = new List<ProductAliasName>();
            _names.Add(new ProductAliasName
            {
                Name = "辐射温度计",
                AliasName = new List<string>() {
                  //  "辐射温度计","辐射温度计（双色）","辐射温度计（光纤）","辐射温度计（光电测温仪）","精密辐射温度计"
                  "辐射温度计","辐射温度计（双色）","辐射温度计（光纤）","辐射温度计（光电测温仪）","精密辐射温度计","红外耳温计","红外体温计（额温型）","辐射温度计（比色）","隐丝式光学高温计"
                }
            });
            _names.Add(new ProductAliasName
            {
                Name = "红外热像仪",
                AliasName = new List<string>() {
                    "红外热像仪"
                }
            });

            _names.Add(new ProductAliasName
            {
                Name = "黑体辐射源",
                AliasName = new List<string>() {
                    "黑体辐射源"
                }
            });
            _names.Add(new ProductAliasName
            {
                Name = "面辐射源",
                AliasName = new List<string>() {
                    "面辐射源","黑体辐射源*"
                }
            });
            _names.Add(new ProductAliasName
            {
                Name = "黑体辐射源_",
                AliasName = new List<string>() {
                    "黑体辐射源_"
                }
            });
        }

        public static string GetProductName(string productName)
        {
            var _aliasName = productName.TrimEnd(' ');

            var list = _names.Where(t => t.AliasName.Contains(_aliasName)).ToList();
            if (list.Count == 0)
                throw new Exception($"未找到产品名称对应的证书生成产品名称.({productName})");
            if (list.Count > 1)
                throw new Exception($"找到多个产品名称对应的证书生成产品名称.({productName})");
            return list[0].Name;

        }









    }
}
