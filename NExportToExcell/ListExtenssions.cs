using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace NExportToExcel
{
    public static class ListExtenssions
    {
        private static List<PropertyInfo> PropertyInfosOf<T, TAttri>()
            where TAttri : Attribute
        {
            return typeof(T).GetProperties()
                .Where(info => info.CustomAttributes.Any(
                    data => data.AttributeType == typeof(TAttri))).ToList();

        }
        public static string ExportToExcel<T>(this List<T> genericList, string path)
            where T : class
        {
            var propertyInfosGenericList = PropertyInfosOf<T, ExportToExcelAttribute>();

            if (propertyInfosGenericList == default)
                throw new Exception("There must be at least one property with the ExportToExcelAttribute tag");

            var columnTitles = GetColumnTitles(propertyInfosGenericList);

            string sheetTitle = SheetTitleOf<T>();
            var mergeBasedOn = MergeColumnBasedOn<T>();

            object valueShouldBeMerged = null;
            int lastOrder = 0;

            using (var excelHelper = new ExcelHelper()
                            .Initialize(path)
                            .AddSheet(sheetTitle)
                            .AddColumnTitles(columnTitles))
            {
                foreach (var currentItem in genericList)
                {
                    excelHelper.GoToTheNextRow();

                    var currentItemProperties = currentItem.GetType().GetProperties();
                    foreach (var propertyInfo in currentItemProperties.Where(info => propertyInfosGenericList.Select(propertyInfo => propertyInfo.Name).Contains(info.Name)))
                    {
                        var value = propertyInfo.GetValue(currentItem);

                        if (propertyInfo.CustomAttributes.Any(a => a.AttributeType == typeof(ShouldBeMergedAttribute)))
                        {
                            if (!string.IsNullOrEmpty(mergeBasedOn))
                            {
                                var orderValue = (int)currentItemProperties.FirstOrDefault(a => a.Name == mergeBasedOn).GetValue(currentItem);

                                if (lastOrder != orderValue || genericList.IndexOf(currentItem) == genericList.Count)
                                {
                                    excelHelper.AddMergeCells(valueShouldBeMerged);

                                    lastOrder = orderValue;
                                }
                                else
                                {
                                    valueShouldBeMerged = value;
                                }
                            }
                        }
                        else
                        {
                            excelHelper.SetCurrentCellValue(value);
                        }
                    }
                }

                excelHelper.Run();

                return excelHelper.FilePath;
            }
        }

        private static string MergeColumnBasedOn<T>() where T : class
        {

            var shouldBeMerged = typeof(T).GetProperties()
                .FirstOrDefault(info => info.CustomAttributes.Any(
                    data => data.AttributeType == typeof(ShouldBeMergedAttribute)));

            return shouldBeMerged != null
              ? shouldBeMerged.CustomAttributes.First(data => data.AttributeType == typeof(ShouldBeMergedAttribute)).ConstructorArguments[0].Value.ToString()
              : string.Empty;

        }

        private static string SheetTitleOf<T>() where T : class
        {
            var attr = typeof(T).CustomAttributes.FirstOrDefault(data => data.AttributeType == typeof(ExportToExcelAttribute));

            var sheetTitle = attr != null && attr.ConstructorArguments.Any()
                             && attr.ConstructorArguments[0].Value.ToString() != null
                ? attr.ConstructorArguments[0].Value.ToString()
                : typeof(T).Name;
            return sheetTitle;
        }

        private static List<string> GetColumnTitles(List<PropertyInfo> propertyInfosGenericList)
        {
            var result = new List<string>();
            foreach (var propertyInfo in propertyInfosGenericList)
            {
                var attribute = propertyInfo.CustomAttributes.FirstOrDefault(data => data.AttributeType == typeof(ExportToExcelAttribute));

                var fieldTitle = "";
                if (attribute != null && attribute.ConstructorArguments.Any()
                    && attribute.ConstructorArguments[0].Value.ToString() != null
                    )
                    fieldTitle = attribute.ConstructorArguments[0].Value.ToString();

                if (string.IsNullOrEmpty(fieldTitle))
                    fieldTitle = propertyInfo.Name;

                result.Add(fieldTitle);
            }

            return result;
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
