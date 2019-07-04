using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class CompareData
    {
        private int _rowIndex;
        private int _similarity;

        public int rowIndex
        {
            get
            {
                return _rowIndex;
            }
        }

        public int similarity
        {
            get
            {
                return _similarity;
            }
        }

        public CompareData(int rowIndex, int similarity)
        {
            _rowIndex = rowIndex;
            _similarity = similarity;
        }
    }
}
