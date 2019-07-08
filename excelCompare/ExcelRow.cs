using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class ExcelRow : IExcelRow
    {
        private ExcelSheet _sheet;
        private int _rowIndex = -1;
        private List<ExcelCell> _columns = new List<ExcelCell>();

        public int rowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public IExcelSheet sheet
        {
            get
            {
                return _sheet;
            }
        }

        public ExcelRow(ExcelSheet sheet, int rowIndex)
        {
            _sheet = sheet;
            _rowIndex = rowIndex;
        }

        internal void AddCell(ExcelCell cell)
        {
            _columns.Add(cell);
            System.Diagnostics.Debug.Assert(_columns.Count <= sheet.columnCount);
        }

        public void GetContent(int columnCount, StringBuilder builder)
        {
            for (int i = 0; i < columnCount; i++)
            {
                if (i > 0)
                {
                    builder.Append(",");
                }
                if (i < _columns.Count)
                {
                    builder.Append(_columns[i]);
                }
                else
                {
                    builder.Append(string.Empty);
                }
            }

            builder.AppendLine();
        }

        public IExcelCell GetCell(int columnIndex)
        {
            if (columnIndex >= 0 && columnIndex < _columns.Count)
            {
                return _columns[columnIndex];
            }
            return null;
        }

        public void BuildContent(StringBuilder builder)
        {
            throw new NotImplementedException();
        }
    }
}
