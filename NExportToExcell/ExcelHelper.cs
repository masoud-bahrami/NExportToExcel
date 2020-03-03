using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace NExportToExcel
{
    /// <summary>
    /// 
    /// </summary>
    public class ExcelHelper : IDisposable
    {
        private Application _excelApp;
        private Workbook _excelWorkBook;
        private int _sheetCounts;
        private Worksheet _excelWorkSheet;
        private string _filePath;
        private string _sheetTitle;

        private int currentCellXIndex = 1;
        private int currentCellYIndex = 1;
        private int _columntCount;
        private bool _shouldSetRowsBackground = true;
        private int numberOfTheFirstRowToBeMerged = 2;

        /// <summary>
        /// 
        /// </summary>
        public ExcelHelper()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ExcelHelper Initialize(string path)
        {
            _filePath = $"{path}\\Book.xlsx";

            _excelApp = new Application();
            _excelApp.Visible = true;
            _excelApp.DefaultSheetDirection = (int)Constants.xlRTL; //or xlRTL
            _excelWorkBook = _excelApp.Workbooks.Add(Type.Missing);

            _sheetCounts = 0;
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        public string FilePath => _filePath;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetTitle"></param>
        /// <returns></returns>
        public ExcelHelper AddSheet(string sheetTitle)
        {
            _sheetTitle = sheetTitle;

            _excelWorkSheet = _excelWorkBook.Sheets.Add();
            _excelWorkSheet.Name = _sheetTitle;
            _sheetCounts++;

            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="firstRow"></param>
        /// <returns></returns>
        public ExcelHelper AddColumnTitles(List<string> firstRow)
        {
            _columntCount = firstRow.Count;

            for (var i = 1; i < firstRow.Count + 1; i++)
            {
                _excelWorkSheet.Cells[1, i] = firstRow[i - 1];
            }

            var columnHeadingsRange = _excelWorkSheet.Range[_excelWorkSheet.Cells[currentCellXIndex, 1], _excelWorkSheet.Cells[currentCellXIndex, _columntCount]];

            columnHeadingsRange.Interior.Color = XlRgbColor.rgbLightGreen;

            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelHelper SetCurrentCellValue(object value)
        {
            _excelWorkSheet.Cells[currentCellXIndex, currentCellYIndex] = value;
            IncrementCurrentCellJIndex();
            return this;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelHelper GoToTheNextRow()
        {
            IncrementCurrentCellXIndex();
            SetCurrentCellJIndex(1);
            return this;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelHelper AddMergeCells(object value)
        {
            _excelWorkSheet.Range[_excelWorkSheet.Cells[numberOfTheFirstRowToBeMerged, _columntCount], _excelWorkSheet.Cells[currentCellXIndex, _columntCount]].Merge();

            _excelWorkSheet.Cells[numberOfTheFirstRowToBeMerged, _columntCount] = value;

            SetBackground();
            UpdateTheNumberOfTheFirstRowToBeMerged();

            return this;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public ExcelHelper Run()
        {
            _filePath = GenerateFileName();
            _excelWorkBook.SaveAs(FilePath);
            _excelWorkBook.Close();
            _excelApp.Quit();
            return this;
        }

        public void Dispose()
        {
            _excelApp = null;
            _excelWorkBook = null;
            _sheetCounts = 0;
            GC.Collect();
        }


        private string GenerateFileName()
        {
            var path = FilePath.Substring(0, FilePath.LastIndexOf("\\") + 1);
            path = path + _sheetTitle + " " + DateTime.Now.ToString("MM-dd-yyyy H-mm-ss") + ".xlsx";
            return path;
        }
        private bool ShouldSetRowsBackground()
        {
            return _shouldSetRowsBackground;
        }
        private void SwapShouldSetRowsBackground()
        {
            _shouldSetRowsBackground = !_shouldSetRowsBackground;
        }
        private void IncrementCurrentCellJIndex()
        {
            ++currentCellYIndex;
        }
        private void SetCurrentCellJIndex(int value)
        {
            currentCellYIndex = value;
        }
        private void IncrementCurrentCellXIndex()
        {
            ++currentCellXIndex;
        }
        private void SetBackground()
        {
            if (ShouldSetRowsBackground())
            {
                var columnHeadingsRange = _excelWorkSheet.Range[_excelWorkSheet.Cells[numberOfTheFirstRowToBeMerged, 1], _excelWorkSheet.Cells[currentCellXIndex, _columntCount]];

                columnHeadingsRange.Interior.Color = XlRgbColor.rgbLightGrey;
            }

            SwapShouldSetRowsBackground();

        }
        private void UpdateTheNumberOfTheFirstRowToBeMerged()
        {
            numberOfTheFirstRowToBeMerged = currentCellXIndex + 1;
        }
    }
}