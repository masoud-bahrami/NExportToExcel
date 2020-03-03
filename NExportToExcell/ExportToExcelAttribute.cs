using System;

namespace NExportToExcel
{
    public class ExportToExcelAttribute : Attribute
    {
        public ExportToExcelAttribute(string excelRowTitle = "")
        {
            this.ExcelRowTitle = excelRowTitle;
        }

        public string ExcelRowTitle { get; }
    }
    public class IgnoreAttribute : Attribute
    {
    }
    public class ShouldBeMergedAttribute : Attribute
    {
        public ShouldBeMergedAttribute(string name)
        {
            MergeBasedOn = name;
        }
        public string MergeBasedOn { get; }
    }
}
