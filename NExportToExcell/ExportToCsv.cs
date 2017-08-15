using System;
using System.Collections.Generic;
using System.Text;

namespace NExportToExcel
{
    internal class ExportToCsv : IDisposable
    {
        private StringBuilder _stringBuilder = new StringBuilder();

        public ExportToCsv InsertFirstRow(List<string> firstRows)
        {
            foreach (var firstRow in firstRows)
            {
                _stringBuilder.Append(FormatCSV(firstRow) + ",");
            }
            _stringBuilder.Append(Environment.NewLine);
            RemoveLastLine();
            return this;
        }

        public ExportToCsv AddBreakeLine()
        {
            _stringBuilder.Append(Environment.NewLine);
            return this;
        }
        public ExportToCsv AddValue(object obj)
        {
            var input = obj != null ? obj.ToString() : string.Empty;
            _stringBuilder.Append(FormatCSV(input) + ",");
            return this;
        }

        public ExportToCsv RemoveLastLine()
        {
            _stringBuilder.Remove(_stringBuilder.Length - 1, 1);
            return this;
        }
        public byte[] Run(string path)
        {
            _stringBuilder.Remove(_stringBuilder.Length - 1, 1);
            _stringBuilder.AppendLine();

            var buffer = Encoding.UTF8.GetBytes(_stringBuilder.ToString());
            return buffer;
        }

        private static string FormatCSV(string input)
        {
            if (input == null)
                return string.Empty;

            var containsQuote = false;
            var containsComma = false;
            var len = input.Length;
            for (var i = 0; i < len && (containsComma == false || containsQuote == false); i++)
            {
                var ch = input[i];
                if (ch == '"')
                {
                    containsQuote = true;
                }
                else if (ch == ',')
                {
                    containsComma = true;
                }
            }

            if (containsQuote && containsComma)
                input = input.Replace("\"", "\"\"");

            if (containsComma)
                return "\"" + input + "\"";
            else
                return input;
        }

        public void Dispose()
        {
            _stringBuilder = null;
            GC.Collect();
        }
    }
}
