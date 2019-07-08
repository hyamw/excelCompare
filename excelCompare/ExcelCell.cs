using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class ExcelCell : IExcelCell
    {
        private int _rowIndex = -1;
        private int _columnIndex = -1;
        private string _value = null;

        public int rowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public int columnIndex
        {
            get
            {
                return _columnIndex;
            }
        }

        public string value
        {
            get
            {
                return _value;
            }
        }

        public string GetContent()
        {
            if (_value == null )
            {
                return string.Empty;
            }
            return _value;
        }

        public ExcelCell(int rowIndex, int columnIndex, string value)
        {
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
            _value = value;
        }

        public ExcelCell(int rowIndex, int columnIndex)
        {
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
        }
    }
}
