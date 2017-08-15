using System;
using System.Collections.Generic;
using System.Linq;

namespace NExportToExcel
{
    public static class ListExtenssions
    {
        public static string ExportToExcel<T>(this List<T> source)
            where T : class
        {
            var propertyInfos1 = typeof(T).GetProperties()
                .Where(info => info.CustomAttributes.Any(
                    data => data.AttributeType == typeof(ExportToExcelAttribute)
                    )).ToList();

            var propertyInfos = new List<string>();
            foreach (var propertyInfo in propertyInfos1)
            {
                var attribute = propertyInfo.CustomAttributes.FirstOrDefault(data => data.AttributeType == typeof(ExportToExcelAttribute));

                var fieldTitle = "";
                if (attribute != null && attribute.ConstructorArguments.Any()
                    && attribute.ConstructorArguments[0].Value.ToString() != null
                    )
                    fieldTitle = attribute.ConstructorArguments[0].Value.ToString();

                if (string.IsNullOrEmpty(fieldTitle))
                    fieldTitle = propertyInfo.Name;

                propertyInfos.Add(fieldTitle);
            }
            var attr = typeof(T).CustomAttributes.FirstOrDefault(data => data.AttributeType == typeof(ExportToExcelAttribute));

            var sheetTitle = attr != null && attr.ConstructorArguments.Any()
                             && attr.ConstructorArguments[0].Value.ToString() != null
                ? attr.ConstructorArguments[0].Value.ToString()
                : typeof(T).Name;

            var excelHelper = new ExcelHelper()
                .Initialize()
                .AddSheet(sheetTitle)
                .AddFirstRow(propertyInfos);

            foreach (var filed in source)
            {
                var properties = filed.GetType().GetProperties();
                var i = 1;
                foreach (var propertyInfo in properties.Where(info => propertyInfos1.Select(propertyInfo => propertyInfo.Name).Contains(info.Name)))
                {
                    var value = propertyInfo.GetValue(filed);
                    if (i == 1)
                        excelHelper.AddFirstRowValue(value);
                    else
                        excelHelper.AddValue(value);
                    i++;

                    Console.WriteLine(value);
                }
            }

            excelHelper.Run();

            excelHelper.Dispose();

            return "";
        }
        
        /// <summary>
        /// Export List of Generic Type to CSV file format, Be notify that only fields with ExportToExcelAttribute
        /// </summary>
        /// <typeparam name="T">Type of List Item</typeparam>
        /// <param name="source">The list which should be converted to csv file</param>
        /// <param name="path">The path, csf file will be saved</param>
        /// <returns></returns>
        public static byte[] ExportToCsv<T>(this List<T> source, string path)
        {
            var propertyInfos1 = typeof(T).GetProperties()
                .Where(info => info.CustomAttributes.Any(
                    data => data.AttributeType == typeof(ExportToExcelAttribute)
                    )).ToList();
            
            var propertyInfos = new List<string>();

            foreach (var propertyInfo in propertyInfos1)
            {
                var attribute = propertyInfo.CustomAttributes.FirstOrDefault(data => data.AttributeType == typeof(ExportToExcelAttribute));

                var fieldTitle = "";
                if (attribute != null && attribute.ConstructorArguments.Any()
                    && attribute.ConstructorArguments[0].Value.ToString() != null
                    )
                    fieldTitle = attribute.ConstructorArguments[0].Value.ToString();

                if (string.IsNullOrEmpty(fieldTitle))
                    fieldTitle = propertyInfo.Name;

                propertyInfos.Add(fieldTitle);
            }

            var addBreakeLine = new ExportToCsv().InsertFirstRow(propertyInfos);

            foreach (var filed in source)
            {
                var properties = filed.GetType().GetProperties();
                foreach (var propertyInfo in properties.Where(info => propertyInfos1.Select(propertyInfo => propertyInfo.Name).Contains(info.Name)))
                {
                    var value = propertyInfo.GetValue(filed);
                    addBreakeLine.AddValue(value);
                }
                addBreakeLine.AddBreakeLine();
                addBreakeLine.RemoveLastLine();
            }
            var stream = addBreakeLine.Run(path);
            addBreakeLine.Dispose();
            return stream;
        }
    }
}
