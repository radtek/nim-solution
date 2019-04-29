using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.Core
{

    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class SimpleExcelMapingAttribute : Attribute
    {
        public SimpleExcelMapingAttribute(string description)
        {
            this.Description = description;
        }

        public string Description
        {
            get; set;
        }
        public string ExcelAddressName
        {
            get; set;
        }
    }
}
