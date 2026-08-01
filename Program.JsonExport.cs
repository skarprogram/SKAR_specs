// SPDX-License-Identifier: GPL-3.0-only

using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;

namespace SKAR_specs
{
    partial class Program
    {
        private static string RenderJsonReport(
            ReportCollectionResult collectionResult,
            string generatedDate,
            StringBuilder errorLog)
        {
            Dictionary<string, object> jsonRoot = new Dictionary<string, object>
            {
                { "format", "SKAR_specs report" },
                { "generatedDate", generatedDate },
                { "collectionError", collectionResult.CollectionException == null ? null : collectionResult.CollectionException.Message },
                { "generationLog", errorLog == null || errorLog.Length == 0 ? "" : errorLog.ToString() },
                { "debugLog", debugEnabled && DebugLogBuilder.Length > 0 ? DebugLogBuilder.ToString() : "" },
                { "report", ToJsonValue(collectionResult.ReportData, true) }
            };

            JavaScriptSerializer serializer = new JavaScriptSerializer
            {
                MaxJsonLength = int.MaxValue,
                RecursionLimit = 100
            };

            return serializer.Serialize(jsonRoot);
        }

        private static object ToJsonValue(object value, bool omitHtmlFields)
        {
            if (value == null)
            {
                return null;
            }

            Type valueType = value.GetType();

            if (value is string ||
                valueType.IsPrimitive ||
                value is decimal ||
                value is DateTime)
            {
                return value;
            }

            if (valueType.IsEnum)
            {
                return value.ToString();
            }

            IDictionary dictionary = value as IDictionary;
            if (dictionary != null)
            {
                Dictionary<string, object> dictionaryResult = new Dictionary<string, object>();
                foreach (DictionaryEntry entry in dictionary)
                {
                    dictionaryResult[entry.Key?.ToString() ?? ""] = ToJsonValue(entry.Value, omitHtmlFields);
                }

                return dictionaryResult;
            }

            IEnumerable enumerable = value as IEnumerable;
            if (enumerable != null)
            {
                List<object> items = new List<object>();
                foreach (object item in enumerable)
                {
                    items.Add(ToJsonValue(item, omitHtmlFields));
                }

                return items;
            }

            Dictionary<string, object> result = new Dictionary<string, object>();
            foreach (PropertyInfo property in valueType.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                if (!property.CanRead || property.GetIndexParameters().Length > 0)
                {
                    continue;
                }

                if (omitHtmlFields && property.Name.EndsWith("Html", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                object propertyValue;
                try
                {
                    propertyValue = property.GetValue(value, null);
                }
                catch
                {
                    continue;
                }

                result[property.Name] = ToJsonValue(propertyValue, omitHtmlFields);
            }

            return result;
        }
    }
}
