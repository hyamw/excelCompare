using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class VRow : IExcelRow
    {
        private VSheet _sheet;
        private int _rowIndex = -1;
        private int _realRowIndex = -1;
        private List<IExcelCell> _columns = new List<IExcelCell>();

        public IExcelSheet sheet
        {
            get
            {
                return _sheet;
            }
        }

        public int rowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public int realRowIndex
        {
            get
            {
                return _realRowIndex;
            }
        }

        public VRow(VSheet sheet, int rowIndex, IExcelRow referenceRow)
        {
            _sheet = sheet;
            _rowIndex = rowIndex;

            if (referenceRow != null)
            {
                if (referenceRow is VRow)
                {
                    VRow row = referenceRow as VRow;
                    _realRowIndex = row.realRowIndex;
                }
                else
                {
                    ExcelRow row = referenceRow as ExcelRow;
                    _realRowIndex = rowIndex;
                }

                for (int columnIndex = 0; columnIndex < sheet.columnCount; columnIndex++)
                {
                    IExcelCell cell = referenceRow.GetCell(columnIndex);
                    IExcelCell newCell = null;
                    if (cell == null)
                    {
                        newCell = new ExcelCell(_rowIndex, columnIndex);
                    }
                    else
                    {
                        newCell = new ExcelCell(_rowIndex, columnIndex, cell.value);
                    }

                    _columns.Add(newCell);
                }
            }
            else
            {
                _realRowIndex = -1;
                for (int columnIndex = 0; columnIndex < sheet.columnCount; columnIndex++)
                {
                    _columns.Add(new ExcelCell(_rowIndex, columnIndex));
                }
            }
        }

        public void BuildContent(StringBuilder builder)
        {
            for (int i = 0; i < sheet.columnCount; i++)
            {
                if (i > 0)
                {
                    builder.Append(",");
                }
                if (i < _columns.Count)
                {
                    builder.Append(_columns[i].GetContent());
                }
                else
                {
                    builder.Append(string.Empty);
                }
            }
        }

        public IExcelCell GetCell(int columnIndex)
        {
            if ( columnIndex >= 0 && columnIndex < _columns.Count )
            {
                return _columns[columnIndex];
            }
            return null;
        }
    }
}
