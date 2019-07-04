using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    public class ExcelRow
    {
        private string _fullText = string.Empty;
        private int _rowIndex = -1;
        private List<string> _columns = new List<string>();
        private int _targetIndex = -1;
        private bool _different = false;

        public List<string> columns
        {
            get
            {
                return _columns;
            }
        }

        public int rowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public int targetIndex
        {
            get
            {
                return _targetIndex;
            }
            set
            {
                _targetIndex = value;
            }
        }

        public bool different
        {
            get
            {
                return _different;
            }
            set
            {
                _different = value;
            }
        }

        public string signature
        {
            get
            {
                return _fullText;
            }
        }

        public ExcelRow(int rowIndex)
        {
            _rowIndex = rowIndex;
        }

        public void Compile()
        {
            StringBuilder builder = new StringBuilder();
            for ( int i = 0; i < columns.Count; i++ )
            {
                if ( i > 0 )
                {
                    builder.Append(",");
                }
                builder.Append(columns[i]);
            }
            _fullText = builder.ToString();
        }

        public string GetColumn(int index)
        {
            if ( index >= 0 && index < _columns.Count )
            {
                return _columns[index];
            }
            return string.Empty;
        }

        public void PaddColumns(int columnCount)
        {
            for ( int i = _columns.Count; i < columnCount; i++ )
            {
                _columns.Add(string.Empty);
            }

            Compile();
        }
    }
}
