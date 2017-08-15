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
}
