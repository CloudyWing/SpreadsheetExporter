using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace CloudyWing.SpreadsheetExporter.Util {
    internal static class DictionaryUtils {
        internal const int MaxNestedPropertyLevel = 5;

        internal static IDictionary<string, object> ConvertFrom<T>(T record, int maxNestedLevel) {
            IDictionary<string, object> dic = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            AddPropertyToDictionary(dic, typeof(T), "", "", record, 0, maxNestedLevel);

            return dic;
        }

        private static void AddPropertyToDictionary(
            IDictionary<string, object> dictionary, Type type, string fulllName, string name, object value, int level, int maxNestedLevel
        ) {
            level++;
            Type underlyingType = Nullable.GetUnderlyingType(type) ?? type;

            if (underlyingType.IsPrimitive || typeof(IEnumerable).IsAssignableFrom(underlyingType)) {
                dictionary.Add(fulllName, value);
            } else {
                if (typeof(IConvertible).IsAssignableFrom(underlyingType)) {
                    dictionary.Add(fulllName, value);
                }

                PropertyInfo[] props = underlyingType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                foreach (PropertyInfo prop in props.Where(x => x.CanRead)) {
                    // DateTime的很多Property還是DateTime，避免無限循環，所以額外判斷
                    if (type == prop.PropertyType && type == typeof(DateTime) && name == prop.Name) {
                        continue;
                    }

                    // 避免類似DateTime情況的class，所以限制層級
                    if (level <= maxNestedLevel) {
                        string _prefix = fulllName.Length == 0 ? prop.Name : $"{fulllName}.{prop.Name}";
                        AddPropertyToDictionary(dictionary, prop.PropertyType, _prefix, prop.Name, value is null ? null : prop.GetValue(value), level, maxNestedLevel);
                    }
                }
            }
        }
    }
}
