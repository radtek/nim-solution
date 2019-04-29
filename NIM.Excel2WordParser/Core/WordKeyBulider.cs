using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace NIM.CertificationGenerator.Core
{
    public class PropertyWordKeyDescription
    {
        public PropertyInfo Property
        {
            get; set;
        }
        public WordKeyAttribute Attribute
        {
            get; set;
        }
    }
    public class WordKeyBulider<T>
    {
        private static List<PropertyWordKeyDescription> propertyAttributes;


        internal static List<PropertyWordKeyDescription> GetPropertyAttributes()
        {
            return propertyAttributes;
        }


        static WordKeyBulider()
        {
            propertyAttributes = _GetPropertyDescriptions();
        }

        private static List<PropertyWordKeyDescription> _GetPropertyDescriptions()
        {
            var propertyInfos = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static).ToList();
            var list = new List<PropertyWordKeyDescription>();
            propertyInfos.ForEach(propertyInfo =>
            {
                if (propertyInfo.CanRead)
                {
                    var attribute = propertyInfo.GetCustomAttributes<WordKeyAttribute>().FirstOrDefault();
                    if (attribute == null)
                        return;
                    if (string.IsNullOrEmpty(attribute.Key))
                        attribute.Key = "{{" + propertyInfo.Name + "}}";
                    list.Add(new PropertyWordKeyDescription
                    {
                        Property = propertyInfo,
                        Attribute = attribute
                    });
                }
            });

            return list;
        }

        public static List<(string replacedString, string replacingString)> GetRepalcedValues(T result)
        {
            var wordkeys = WordKeyBulider<T>.GetPropertyAttributes();

            var list = new List<(string replacedString, string replacingString)>();
            wordkeys.ForEach(t =>
            {
                var replacedString = t.Attribute.Key;
                var _value = t.Property.GetValue(result);
                var replacingString = "";
                if (_value == null)
                    replacingString = "";
                else
                    replacingString = _value.ToString();

                list.Add((replacedString, replacingString));

            });
            return list;
        }


    }
}
