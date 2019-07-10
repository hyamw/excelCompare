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
        private string _content = null;

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
            if (_content == null )
            {
                return string.Empty;
            }
            return _content;
        }

        public ExcelCell(int rowIndex, int columnIndex, string value)
        {
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
            _value = value;
            _content = value;
            if ( !string.IsNullOrEmpty(_content) )
            {
                _content = _content.Replace("\r\n", SheetComparer.RETURN_NEWLINE_PLACE_HOLDER);
                _content = _content.Replace("\r", SheetComparer.RETURN_PLACE_HOLDER);
                _content = _content.Replace("\n", SheetComparer.NEWLINE_PLACE_HOLDER);
            }
        }

        public ExcelCell(int rowIndex, int columnIndex)
        {
            _rowIndex = rowIndex;
            _columnIndex = columnIndex;
        }
    }
}
