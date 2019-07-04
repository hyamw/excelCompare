using DiffMatchPatch;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class SheetComparer
    {
        private string _name;
        private ExcelSheet _left;
        private ExcelSheet _right;
        private int _columnCount;
        private int _rowCount;
        private bool _isDifferent = false;
        private bool[,] _cellStatus = null;
        private bool[] _rowStatus = null;

        public string name
        {
            get
            {
                return _name;
            }
        }

        public ExcelSheet left
        {
            get
            {
                return _left;
            }
        }

        public ExcelSheet right
        {
            get
            {
                return _right;
            }
        }

        public int rowCount
        {
            get
            {
                return _rowCount;
            }
        }

        public int columnCount
        {
            get
            {
                return _columnCount;
            }
        }

        public bool isDifferent
        {
            get
            {
                return _isDifferent;
            }
        }

        public SheetComparer(string name)
        {
            _name = name;
        }

        public void Execute(ExcelSheet left, ExcelSheet right)
        {
            _left = left;
            _right = right;

            _columnCount = Math.Max(left.columnCount, right.columnCount);
            left.PaddColumns(columnCount);
            right.PaddColumns(columnCount);

            diff_match_patch comparer = new diff_match_patch();
            string leftContent = left.GetContent();
            string rightContent = right.GetContent();
            List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
            comparer.diff_cleanupSemanticLossless(diffs);

            DiffAnalyzer analyzer = new DiffAnalyzer();
            _isDifferent = analyzer.Analysis(diffs, left, right);

            _rowCount = left.diffViewRowCount;

            _cellStatus = new bool[_rowCount, _columnCount];
            _rowStatus = new bool[_rowCount];
            for ( int rowIndex = 0; rowIndex < _rowCount; rowIndex++ )
            {
                _rowStatus[rowIndex] = false;
                for (int columnIndex = 0; columnIndex < _columnCount; columnIndex++)
                {
                    _cellStatus[rowIndex, columnIndex] = left.CompareCell(right, rowIndex, columnIndex);
                    if (_cellStatus[rowIndex, columnIndex])
                    {
                        _rowStatus[rowIndex] = true;
                    }
                }
            }
        }

        public void Execute(ExcelWorkbook left, ExcelWorkbook right)
        {
            Execute(left.LoadSheet(name), right.LoadSheet(name));
        }

        public static void Compare(ExcelSheet left, ExcelSheet right)
        {
            int diffColumnCount = Math.Max(left.columnCount, right.columnCount);
            left.PaddColumns(diffColumnCount);
            right.PaddColumns(diffColumnCount);
            diff_match_patch comparer = new diff_match_patch();
            string leftContent = left.GetContent();
            string rightContent = right.GetContent();
            List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
            comparer.diff_cleanupSemanticLossless(diffs);

            DiffAnalyzer analyzer = new DiffAnalyzer();
            analyzer.Analysis(diffs, left, right);

            string l = analyzer.GetLeftContent();
            string r = analyzer.GetRightContent();

            if ( l.Equals(leftContent) && r.Equals(rightContent) )
            {
                int a = 0;
                a++;
            }
        }

        public bool IsRowDifferent(int rowIndex)
        {
            if ( rowIndex >= 0 && rowIndex < rowCount )
            {
                return _rowStatus[rowIndex];
            }
            return false;
        }

        public bool IsCellDifferent(int rowIndex, int columnIndex)
        {
            if (rowIndex >= 0 && rowIndex < rowCount && columnIndex >= 0 && columnIndex < columnCount )
            {
                return _cellStatus[rowIndex, columnIndex];
            }
            return false;
        }
    }
}
