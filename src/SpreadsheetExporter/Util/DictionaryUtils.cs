using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace CloudyWing.SpreadsheetExporter.Util {
    internal static class DictionaryUtils {
        internal static IDictionary<string, object> ConvertFrom<T>(T record, int maxNestedLevel) {
            Dictionary<string, object> dictionary = new(StringComparer.OrdinalIgnoreCase);
            AddPropertyToDictionary(dictionary, "", record, 0, maxNestedLevel);
            return dictionary;
        }


        private static void AddPropertyToDictionary(
            IDictionary<string, object> dictionary, string fullName, object value, int level, int maxNestedLevel) {
            if (value is null) {
                dictionary.Add(fullName, null);
                return;
            }

            Type type = value.GetType();
            Type underlyingType = Nullable.GetUnderlyingType(type) ?? type;

            if (ShouldStoreDirectly(underlyingType)) {
                dictionary.Add(fullName, value);
                return;
            }

            if (level >= maxNestedLevel) {
                dictionary.Add(fullName, value);
                return;
            }

            IEnumerable<PropertyInfo> properties = underlyingType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead);

            foreach (PropertyInfo prop in properties) {
                object propValue = prop.GetValue(value);
                string propFullName = string.IsNullOrEmpty(fullName)
                    ? prop.Name
                    : $"{fullName}.{prop.Name}";

                AddPropertyToDictionary(
                    dictionary,
                    propFullName,
                    propValue,
                    level + 1,
                    maxNestedLevel
                );
            }
        }

        private static bool ShouldStoreDirectly(Type type) {
            return type.IsPrimitive
                || typeof(IEnumerable).IsAssignableFrom(type)
                || typeof(IConvertible).IsAssignableFrom(type)
                || type == typeof(DateTimeOffset)
                || type == typeof(TimeSpan)
                || type == typeof(Guid);
        }
    }
}
