using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace CloudyWing.SpreadsheetExporter.Util {
    internal static class DictionaryUtils {
        internal static IDictionary<string, object> ConvertFrom<T>(T valueObj) {
            IDictionary<string, object> dic = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            AddPropertyToDictionary(typeof(T), "", valueObj);

            return dic;

            void AddPropertyToDictionary(Type _type, string _name, object _value) {
                Type underlyingType = Nullable.GetUnderlyingType(_type) ?? _type;

                if (underlyingType.IsPrimitive
                    || typeof(IConvertible).IsAssignableFrom(underlyingType)
                    || typeof(IEnumerable).IsAssignableFrom(underlyingType) && underlyingType != typeof(string)
                ) {
                    dic.Add(_name, _value);
                } else {
                    PropertyInfo[] props = underlyingType.GetProperties(BindingFlags.Public | BindingFlags.DeclaredOnly | BindingFlags.Instance);
                    foreach (PropertyInfo prop in props.Where(x => x.CanRead)) {
                        string _prefix = _name.Length == 0 ? prop.Name : $"{_name}.{prop.Name}";
                        AddPropertyToDictionary(prop.PropertyType, _prefix, _value is null ? null : prop.GetValue(_value));
                    }
                }
            }
        }
    }
}
