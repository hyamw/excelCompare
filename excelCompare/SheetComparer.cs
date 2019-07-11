using DiffMatchPatch;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class SheetComparer
    {
        public const string LINE_SPLITTER = "=[$$]=";

        public const string RETURN_PLACE_HOLDER = "=[@@]=";
        public const string NEWLINE_PLACE_HOLDER = "=[**]=";
        public const string RETURN_NEWLINE_PLACE_HOLDER = "=[++]=";

        private static readonly string[] SPLIT_SEPERATORS = new string[] { LINE_SPLITTER };
        private Content leftContent = new Content();
        private Content rightContent = new Content();

        private VSheet _left;
        private VSheet _right;
        private int _columnCount;
        private bool _isDifferent = false;
        private List<RowComparer> _rows = new List<RowComparer>();
        private int _leftAlignIndex = -1;
        private int _rightAlignIndex = -1;

        public VSheet left
        {
            get
            {
                return _left;
            }
        }

        public VSheet right
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
                return _rows.Count;
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

        public void Execute(IExcelSheet left, IExcelSheet right, int leftAlignIndex = -1, int rightAlignIndex = -1)
        {
            _isDifferent = false;

            _columnCount = Math.Max(left.columnCount, right.columnCount);
            _left = new VSheet(left, _columnCount);
            _right = new VSheet(right, _columnCount);

            _leftAlignIndex = leftAlignIndex;
            _rightAlignIndex = rightAlignIndex;
            if (_leftAlignIndex != -1 && _rightAlignIndex != -1)
            {
                int beginRow = 0;
                int endRow = 0;
                if (_leftAlignIndex > _rightAlignIndex)
                {
                    endRow = _leftAlignIndex - 1;
                    ((VSheet)_right).PadRowsAt(_rightAlignIndex, (_leftAlignIndex - _rightAlignIndex));
                }
                else
                {
                    endRow = _rightAlignIndex - 1;
                    ((VSheet)_left).PadRowsAt(_leftAlignIndex, (_rightAlignIndex - _leftAlignIndex));
                }

                leftContent.Clear();
                rightContent.Clear();
                {
                    diff_match_patch comparer = new diff_match_patch();
                    string leftContent = _left.GetContent(beginRow, endRow);
                    string rightContent = _right.GetContent(beginRow, endRow);
                    List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
                    comparer.diff_cleanupSemanticLossless(diffs);

                    Compare(diffs, beginRow, beginRow);
                }

                leftContent.Clear();
                rightContent.Clear();
                {
                    endRow++;
                    diff_match_patch comparer = new diff_match_patch();
                    string leftContent = _left.GetContent(endRow, endRow);
                    string rightContent = _right.GetContent(endRow, endRow);
                    List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
                    comparer.diff_cleanupSemanticLossless(diffs);

                    Compare(diffs, endRow, endRow);
                }

                leftContent.Clear();
                rightContent.Clear();
                {
                    endRow++;
                    diff_match_patch comparer = new diff_match_patch();
                    string leftContent = _left.GetContent(endRow, _left.rowCount - 1);
                    string rightContent = _right.GetContent(endRow, _right.rowCount - 1);
                    List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
                    comparer.diff_cleanupSemanticLossless(diffs);

                    Compare(diffs, endRow, endRow);
                }
            }
            else
            {

                diff_match_patch comparer = new diff_match_patch();
                string leftContent = _left.GetContent();
                string rightContent = _right.GetContent();
                List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
                comparer.diff_cleanupSemanticLossless(diffs);

                Compare(diffs, 0, 0);
            }

            VSheet leftResult = new VSheet(_left.name, _columnCount);
            VSheet rightResult = new VSheet(_right.name, _columnCount);

            List<int> deleteList = new List<int>();
            for ( int i = 0; i < _rows.Count; i++ )
            {
                RowComparer row = _rows[i];
                IExcelRow leftRow = _left.GetRow(row.leftRowIndex);
                IExcelRow rightRow = _right.GetRow(row.rightRowIndex);
                if (GetRealRowIndex(leftRow) == -1 && GetRealRowIndex(rightRow) == -1)
                {
                    deleteList.Add(i);
                    continue;
                }
                leftResult.NewRow(leftRow);
                rightResult.NewRow(rightRow);
            }

            for ( int i = deleteList.Count - 1; i >= 0; i-- )
            {
                _rows.RemoveAt(deleteList[i]);
            }

            _left = leftResult;
            _right = rightResult;
        }

        private int GetRealRowIndex(IExcelRow row)
        {
            int index = -1;
            if ( row != null && row is VRow )
            {
                index = ((VRow)row).realRowIndex;
            }
            return index;
        }

        public void Compare(List<Diff> diffs, int leftBeginRowIndex, int rightBeginRowIndex)
        {
            for (int diffIndex = 0; diffIndex < diffs.Count; diffIndex++)
            {
                Diff diff = diffs[diffIndex];
                string[] lines = diff.text.Split(SPLIT_SEPERATORS, StringSplitOptions.None);
                for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
                {
                    string line = lines[lineIndex];

                    if (lineIndex > 0)
                    {
                        if (diff.operation == Operation.EQUAL)
                        {
                            bool leftDiff = leftContent.FinishLine();
                            bool rightDiff = rightContent.FinishLine();
                            AddRow(leftBeginRowIndex + leftContent.LineCount - 1, rightBeginRowIndex + rightContent.LineCount - 1, leftDiff || rightDiff);
                        }
                        else if (diff.operation == Operation.DELETE)
                        {
                            leftContent.FinishLine();
                            AddRow(leftBeginRowIndex + leftContent.LineCount - 1, -1, true);
                        }
                        else if (diff.operation == Operation.INSERT)
                        {
                            rightContent.FinishLine();
                            AddRow(-1, rightBeginRowIndex + rightContent.LineCount - 1, true);
                        }
                    }

                    if (diff.operation == Operation.EQUAL)
                    {
                        leftContent.EnqueueLine(line, false);
                        rightContent.EnqueueLine(line, false);
                    }
                    else if (diff.operation == Operation.DELETE)
                    {
                        leftContent.EnqueueLine(line, !string.IsNullOrEmpty(line));
                        _isDifferent = true;
                    }
                    else if (diff.operation == Operation.INSERT)
                    {
                        rightContent.EnqueueLine(line, !string.IsNullOrEmpty(line));
                        _isDifferent = true;
                    }
                }
            }
        }

        public bool IsRowDifferent(int rowIndex)
        {
            if ( rowIndex >= 0 && rowIndex < rowCount )
            {
                return _rows[rowIndex].isDifferent;
            }
            return false;
        }

        public bool IsCellDifferent(int rowIndex, int columnIndex)
        {
            if (rowIndex >= 0 && rowIndex < rowCount && columnIndex >= 0 && columnIndex < columnCount )
            {
                return _rows[rowIndex].IsCellDifferent(columnIndex);
            }
            return false;
        }

        private void AddRow(int leftRowIndex, int rightRowIndex, bool isDifferent)
        {
            RowComparer newRow = new RowComparer(_rows.Count, leftRowIndex, rightRowIndex, isDifferent);
            _rows.Add(newRow);
            newRow.CompareCells(_columnCount, left, right);
        }

        public void Copy2Left(int rowIndex)
        {
            VRow leftRow = _left.GetRow(rowIndex) as VRow;
            VRow rightRow = _right.GetRow(rowIndex) as VRow;
            leftRow.CopyFrom(rightRow);
            _rows[rowIndex] = new RowComparer(rowIndex, rowIndex, rowIndex, false);
            _rows[rowIndex].CompareCells(_columnCount, left, right);
        }

        public void Copy2Right(int rowIndex)
        {
            VRow leftRow = _left.GetRow(rowIndex) as VRow;
            VRow rightRow = _right.GetRow(rowIndex) as VRow;
            rightRow.CopyFrom(leftRow);
            _rows[rowIndex] = new RowComparer(rowIndex, rowIndex, rowIndex, false);
            _rows[rowIndex].CompareCells(_columnCount, left, right);
        }

        protected class Content
        {
            private string openEntry = null;
            private List<string> lines = new List<string>();
            private bool openDiff = false;

            public void Clear()
            {
                openEntry = null;
                lines.Clear();
                openDiff = false;
            }

            public void AddLine(string line)
            {
                lines.Add(line);
            }

            public void EnqueueLine(string text, bool different)
            {
                if (openEntry == null)
                {
                    openEntry = text;
                }
                else
                {
                    openEntry += text;
                }

                openDiff |= different;
            }

            public void FinishLine(string text)
            {
                if (openEntry != null)
                {
                    if (text != null)
                    {
                        lines.Add(openEntry + text);
                    }
                    else
                    {
                        lines.Add(openEntry);
                    }
                }
                else
                {
                    if (text != null)
                    {
                        lines.Add(text);
                    }
                    else
                    {
                        Debug.Assert(false, "Unexpected!");
                    }
                }
                openEntry = null;
                openDiff = false;
            }

            public bool FinishLine()
            {
                bool result = openDiff;
                if (openEntry != null)
                {
                    lines.Add(openEntry);
                }
                else
                {
                    Debug.Assert(false, "Unexpected!");
                }
                openEntry = null;
                openDiff = false;
                return result;
            }

            public string GetFullContent()
            {
                StringBuilder builder = new StringBuilder();
                for (int i = 0; i < lines.Count; i++)
                {
                    builder.AppendLine(lines[i]);
                }
                return builder.ToString();
            }

            public int LineCount
            {
                get
                {
                    return lines.Count;
                }
            }
        }
    }
}
