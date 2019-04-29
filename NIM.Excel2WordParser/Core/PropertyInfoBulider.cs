using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.Core
{

    public class PropertyDescription
    {
        public PropertyInfo Property
        {
            get; set;
        }
        public SimpleExcelMapingAttribute Attribute
        {
            get; set;
        }
    }
    public class PropertyInfoBulider<T>
    {
        private static List<PropertyDescription> propertyAttributes;
        

        internal static List<PropertyDescription> GetPropertyAttributes()
        {
            return propertyAttributes;
        }


        static PropertyInfoBulider()
        {
            propertyAttributes = _GetPropertyDescriptions();
        }

        private static List<PropertyDescription> _GetPropertyDescriptions()
        {
            var propertyInfos = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static).ToList();
            var list = new List<PropertyDescription>();
            propertyInfos.ForEach(propertyInfo =>
            {
                if (propertyInfo.CanWrite && propertyInfo.CanWrite)
                {
                    var attribute = propertyInfo.GetCustomAttributes<SimpleExcelMapingAttribute>().FirstOrDefault();
                    if (attribute == null)
                        return;
                    list.Add(new PropertyDescription
                    {
                        Property = propertyInfo,
                        Attribute = attribute
                    });
                }
            });
      

            return list;
        }



    }
}
