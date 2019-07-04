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

        private static readonly string[] SPLIT_SEPERATORS = new string[] { LINE_SPLITTER };
        private Content leftContent = new Content();
        private Content rightContent = new Content();

        private string _name;
        private ExcelSheet _left;
        private ExcelSheet _right;
        private int _columnCount;
        private bool _isDifferent = false;
        private List<RowComparer> _rows = new List<RowComparer>();

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

        public List<RowComparer> rows
        {
            get
            {
                return _rows;
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

            diff_match_patch comparer = new diff_match_patch();
            string leftContent = left.GetContent(_columnCount);
            string rightContent = right.GetContent(_columnCount);
            List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
            comparer.diff_cleanupSemanticLossless(diffs);

            Compare(diffs);
        }

        public void Execute(ExcelWorkbook left, ExcelWorkbook right)
        {
            Execute(left.LoadSheet(name), right.LoadSheet(name));
        }

        public void Compare(List<Diff> diffs)
        {
            _isDifferent = false;
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
                            AddRow(leftContent.LineCount - 1, rightContent.LineCount - 1, leftDiff || rightDiff);
                        }
                        else if (diff.operation == Operation.DELETE)
                        {
                            leftContent.FinishLine();
                            AddRow(leftContent.LineCount - 1, -1, true);
                        }
                        else if (diff.operation == Operation.INSERT)
                        {
                            rightContent.FinishLine();
                            AddRow(-1, rightContent.LineCount - 1, true);
                        }
                    }

                    if (diff.operation == Operation.EQUAL)
                    {
                        leftContent.EnqueueLine(line, false);
                        rightContent.EnqueueLine(line, false);
                    }
                    else if (diff.operation == Operation.DELETE)
                    {
                        leftContent.EnqueueLine(line, true);
                        _isDifferent = true;
                    }
                    else if (diff.operation == Operation.INSERT)
                    {
                        rightContent.EnqueueLine(line, true);
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

        public DataTable GetLeftSource()
        {
            DataTable dt = new DataTable();
            for (int j = 0; j < _columnCount; j++)
            {
                dt.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(j)));
            }

            int rowIndex = 0;
            DataRow dr = null;
            for (int i = 0; i < rowCount; i++)
            {
                dr = dt.NewRow();
                if (_rows[i].leftRowIndex != -1)
                {
                    ExcelRow row = left.GetRowByRealIndex(_rows[i].leftRowIndex);
                    for (int k = 0; k < _columnCount; k++)
                    {
                        if (k < columnCount && row != null)
                        {
                            dr[k] = row.GetColumn(k);
                        }
                        else
                        {
                            dr[k] = string.Empty;
                        }
                    }

                    rowIndex++;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public DataTable GetRightSource()
        {
            DataTable dt = new DataTable();
            for (int j = 0; j < _columnCount; j++)
            {
                dt.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(j)));
            }

            int rowIndex = 0;
            DataRow dr = null;
            for (int i = 0; i < rowCount; i++)
            {
                dr = dt.NewRow();
                if (_rows[i].rightRowIndex != -1)
                {
                    ExcelRow row = right.GetRowByRealIndex(_rows[i].rightRowIndex);
                    for (int k = 0; k < _columnCount; k++)
                    {
                        if (k < columnCount && row != null)
                        {
                            dr[k] = row.GetColumn(k);
                        }
                        else
                        {
                            dr[k] = string.Empty;
                        }
                    }

                    rowIndex++;
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        private void AddRow(int leftRowIndex, int rightRowIndex, bool isDifferent)
        {
            RowComparer newRow = new RowComparer(_rows.Count, leftRowIndex, rightRowIndex, isDifferent);
            _rows.Add(newRow);
            newRow.CompareCells(_columnCount, left, right);
        }

        protected class Content
        {
            private string openEntry = null;
            private List<string> lines = new List<string>();
            private bool openDiff = false;

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

        public ExcelRow GetLeftRow(int index)
        {
            return left.GetRowByRealIndex(_rows[index].leftRowIndex);
        }

        public ExcelRow GetRightRow(int index)
        {
            return right.GetRowByRealIndex(_rows[index].rightRowIndex);
        }
    }
}
