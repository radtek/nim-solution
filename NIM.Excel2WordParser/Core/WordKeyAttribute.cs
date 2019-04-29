using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.Core
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class WordKeyAttribute : Attribute
    {
        public WordKeyAttribute()
        {

        }
        public WordKeyAttribute(string key)
        {
            this.Key = key;
        }
        public string Key
        {
            get;set;
        }
    }
}
