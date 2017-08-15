using System;
using System.Collections.Generic;

namespace NExportToExcel
{
   /// <summary>
   /// 
   /// </summary>
   public class ExcelHelper : IDisposable
    {
        private Microsoft.Office.Interop.Excel.Application  _excelApp;
        private Microsoft.Office.Interop.Excel.Workbook _excelWorkBook;
        private int _sheetCounts;
        private Microsoft.Office.Interop.Excel.Worksheet _excelWorkSheet;
        private string _filePath;
        /// <summary>
        /// 
        /// </summary>
        public ExcelHelper()
        {
            Initialize();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ExcelHelper Initialize()
        {
            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelWorkBook = _excelApp.Workbooks.Add(Type.Missing);
            _sheetCounts = 0;
            _filePath = "";
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        public string FilePath => _filePath;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetTitles"></param>
        /// <returns></returns>
        public ExcelHelper AddSheets(List<string> sheetTitles)
        {
            foreach (var sheetTitle in sheetTitles)
            {
                AddSheet(sheetTitle);
                _sheetCounts++;
            }
            return this;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetTitle"></param>
        /// <returns></returns>
        public ExcelHelper AddSheet(string sheetTitle)
        {
            _excelWorkSheet = _excelWorkBook.Sheets.Add();
            _excelWorkSheet.Name = sheetTitle;
            _sheetCounts++;
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="firstRow"></param>
        /// <returns></returns>
        public ExcelHelper AddFirstRow(List<string> firstRow)
        {
            for (var i = 1; i < firstRow.Count + 1; i++)
            {
                _excelWorkSheet.Cells[1, i] = firstRow[i - 1];
            }
            return this;
        }

        private static int i = 1;
        private static int j = 2;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelHelper AddFirstRowValue(object value)
        {
            i++;
            _excelWorkSheet.Cells[i, 1] = value;
            j = 2;
            return this;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelHelper AddValue(object value)
        {
            _excelWorkSheet.Cells[i, j] = value;
            j++;
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ExcelHelper Run()
        {
            _excelWorkBook.Save();
            _excelWorkBook.Close();
            _excelApp.Quit();
            return this;
        }
        public void Dispose()
        {
            _excelApp = null;
            _excelWorkBook = null;
            _sheetCounts = 0;
            _filePath = string.Empty;
            GC.Collect();
        }
    }
}
