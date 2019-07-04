using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace excelCompare
{
    public partial class MainForm : Form
    {
        private Color DIFF_BACK_COLOR = Color.FromArgb(255, 227, 227);
        private Color DIFF_FORE_COLOR = Color.FromArgb(255, 0, 0);
        private bool stopUpdate = false;
        private DataTable lineTable = new DataTable();
        private ExcelComparer comparer = new ExcelComparer();
        private SheetComparer currentSheet = null;

        public MainForm(string[] args)
        {
            InitializeComponent();

            if ( args != null && args.Length >= 2 )
            {
                if (args.Length == 3)
                {
                    if (!File.Exists(args[2]))
                    {
                        MessageBox.Show(this, "合并结果文件不存在!");
                        return;
                    }
                    else
                    {
                        MessageBox.Show(this, "目前尚不支持合并功能!");
                    }
                }

                bool leftExists = File.Exists(args[0]);
                bool rightExists = File.Exists(args[1]);
                if (leftExists && rightExists)
                {
                    CompareExcel(args[0], args[1]);
                }
                else
                {
                    if (!leftExists)
                    {
                        MessageBox.Show(this, string.Format("对比文件:{0}不存在!", args[0]));
                    }
                    else if ( !rightExists )
                    {
                        MessageBox.Show(this, string.Format("对比文件:{0}不存在!", args[1]));
                    }
                }
            }
        }

        private void OnCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (currentSheet == null)
            {
                return;
            }

            if (currentSheet.IsRowDifferent(e.RowIndex))
            {
                e.CellStyle.BackColor = DIFF_BACK_COLOR;
                if (currentSheet.IsCellDifferent(e.RowIndex, e.ColumnIndex))
                {
                    e.CellStyle.ForeColor = DIFF_FORE_COLOR;
                }
                else
                {
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            else
            {
                e.CellStyle.BackColor = Color.White;
                e.CellStyle.ForeColor = Color.Black;
            }
        }

        private DataGridView GetSource(object sender)
        {
            return sender as DataGridView;
        }

        private DataGridView GetTarget(object sender)
        {
            if (sender == leftGrid)
            {
                return rightGrid;
            }
            else
            {
                return leftGrid;
            }
        }

        private void OnColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            if (stopUpdate)
            {
                return;
            }
            DataGridView target = GetTarget(sender);

            target.Columns[e.Column.Index].Width = e.Column.Width;
        }

        private void OnRowHeaderWidthChanged(object sender, EventArgs e)
        {
            if (stopUpdate)
            {
                return;
            }
            DataGridView source = GetSource(sender);
            DataGridView target = GetTarget(sender);
            target.RowHeadersWidth = source.RowHeadersWidth;
        }

        private void OnRowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if (e.Row.Index >= 0 && e.Row.Index < currentSheet.rowCount)
            {
                if (sender == leftGrid)
                {
                    int realIndex = currentSheet.rows[e.Row.Index].leftRowIndex;
                    if (realIndex != -1)
                    {
                        e.Row.HeaderCell.Value = (realIndex + 1).ToString();
                    }
                    else
                    {
                        e.Row.HeaderCell.Value = string.Empty;
                    }
                }
                else
                {
                    int realIndex = currentSheet.rows[e.Row.Index].rightRowIndex;
                    if (realIndex != -1)
                    {
                        e.Row.HeaderCell.Value = (realIndex + 1).ToString();
                    }
                    else
                    {
                        e.Row.HeaderCell.Value = string.Empty;
                    }
                }
            }
        }

        private void OnGridViewScroll(object sender, ScrollEventArgs e)
        {
            DataGridView source = GetSource(sender);
            DataGridView target = GetTarget(sender);
            if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
            {
                target.FirstDisplayedScrollingRowIndex = source.FirstDisplayedScrollingRowIndex;
            }
            else
            {
                target.HorizontalScrollingOffset = source.HorizontalScrollingOffset;
            }
        }

        private DataGridView updatingGrid = null;
        private void OnGridViewSelectionChanged(object sender, EventArgs e)
        {
            if (stopUpdate || sender == updatingGrid )
            {
                return;
            }
            DataGridView source = GetSource(sender);
            DataGridView target = GetTarget(sender);
            if ( source.SelectedCells.Count > 0 )
            {
                updatingGrid = target;
                target.ClearSelection();
                target.Rows[source.SelectedCells[0].RowIndex].Cells[source.SelectedCells[0].ColumnIndex].Selected = true;
                updatingGrid = null;

                ShowLineDifference(source.SelectedCells[0].RowIndex);
                UpdateNextPreviousButton();
            }
        }

        private void OnSheetSelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSheet(comparer.sheetNames[sheetNamesComboBox.SelectedIndex]);
        }

        private void LoadSheet(string name)
        {
            currentSheet = comparer.GetSheet(name);
            stopUpdate = true;

            leftGrid.DataSource = currentSheet.GetLeftSource();
            rightGrid.DataSource = currentSheet.GetRightSource();
            stopUpdate = false;

            lineTable = new DataTable();

            for ( int i = 0; i < currentSheet.columnCount; i++ )
            {
                lineTable.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(i)));
            }

            rowDiffGrid.DataSource = lineTable;

            UpdateNextPreviousButton();
        }

        private void OnPreviousButtonClicked(object sender, EventArgs e)
        {
            SelectPreviousDifference();
        }

        private void SelectPreviousDifference()
        {
            int rowIndex = GetPreviousDifferenctRowIndex();
            if (rowIndex != -1)
            {
                JumToRow(rowIndex);
            }
        }

        private void OnNextButtonClicked(object sender, EventArgs e)
        {
            SelectNextDifference();
        }

        private void SelectNextDifference()
        {
            int rowIndex = GetNextDifferenctRowIndex();
            if (rowIndex != -1)
            {
                JumToRow(rowIndex);
            }
        }

        private void JumToRow(int rowIndex)
        {
            stopUpdate = true;
            leftGrid.ClearSelection();
            leftGrid.Rows[rowIndex].Selected = true;
            rightGrid.Rows[rowIndex].Selected = true;
            leftGrid.CurrentCell = leftGrid.Rows[rowIndex].Cells[0];
            rightGrid.CurrentCell = rightGrid.Rows[rowIndex].Cells[0];
            stopUpdate = false;
            ShowLineDifference(rowIndex);

            UpdateNextPreviousButton();
        }

        int GetCurrentRowIndex()
        {
            int currentRowIndex = 0;
            if (leftGrid.SelectedCells.Count > 0)
            {
                currentRowIndex = leftGrid.SelectedCells[0].RowIndex;
            }
            return currentRowIndex;
        }

        private int GetNextDifferenctRowIndex()
        {
            int currentRowIndex = GetCurrentRowIndex();
            for (int i = currentRowIndex + 1; i < currentSheet.rowCount; i++)
            {
                if (currentSheet.IsRowDifferent(i))
                {
                    return i;
                }
            }
            return -1;
        }

        private int GetPreviousDifferenctRowIndex()
        {
            int currentRowIndex = GetCurrentRowIndex();
            for (int i = currentRowIndex - 1; i >= 0; i--)
            {
                if (currentSheet.IsRowDifferent(i))
                {
                    return i;
                }
            }
            return -1;
        }

        private string SafeGetColumn(ExcelRow row, int columnIndex)
        {
            if ( row == null )
            {
                return string.Empty;
            }
            if ( columnIndex < row.columns.Count )
            {
                return row.columns[columnIndex];
            }
            return string.Empty;
        }

        private void ShowLineDifference(int rowIndex)
        {
            rowDiffGrid.DataSource = null;

            ExcelRow leftRow = currentSheet.GetLeftRow(rowIndex);
            ExcelRow rightRow = currentSheet.GetRightRow(rowIndex);
            int columnCount = currentSheet.columnCount;
            lineTable.Rows.Clear();
            DataRow row = lineTable.NewRow();
            for ( int i = 0; i < columnCount; i++ )
            {
                row[i] = SafeGetColumn(leftRow, i);
            }
            lineTable.Rows.Add(row);

            row = lineTable.NewRow();
            for (int i = 0; i < columnCount; i++)
            {
                row[i] = SafeGetColumn(rightRow, i);
            }
            lineTable.Rows.Add(row);

            rowDiffGrid.DataSource = lineTable;
        }

        private bool CompareColumn(int columnIndex)
        {
            DataRow firstRow = lineTable.Rows[0];
            DataRow secondRow = lineTable.Rows[1];
            return !firstRow[columnIndex].Equals(secondRow[columnIndex]);
        }

        private bool CompareRow()
        {
            if (lineTable.Rows.Count < 2)
            {
                return true;
            }
            DataRow firstRow = lineTable.Rows[0];
            DataRow secondRow = lineTable.Rows[1];
            for (int i = 0; i < lineTable.Columns.Count; i++)
            {
                if ( !firstRow[i].Equals(secondRow[i]) )
                {
                    return false;
                }
            }
            return true;
        }

        private void OnRowGridCellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (!CompareRow())
            {
                e.CellStyle.BackColor = DIFF_BACK_COLOR;
                if (CompareColumn(e.ColumnIndex))
                {
                    e.CellStyle.ForeColor = DIFF_FORE_COLOR;
                }
                else
                {
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            else
            {
                e.CellStyle.BackColor = Color.White;
                e.CellStyle.ForeColor = Color.Black;
            }
        }

        private void CompareExcel(string leftPath, string rightPath)
        {
            comparer.Load(leftPath, rightPath);

            sheetNamesComboBox.Items.Clear();

            for (int i = 0; i < comparer.sheetNames.Length; i++)
            {
                string sheetName = comparer.sheetNames[i];
                string differentMark = comparer.GetSheet(sheetName).isDifferent ? "[*]" : string.Empty;
                sheetNamesComboBox.Items.Add(sheetName + differentMark);
            }

            sheetNamesComboBox.SelectedIndex = 0;

            LoadSheet(comparer.sheetNames[0]);
        }

        private void OnOpenClicked(object sender, EventArgs e)
        {
            OpenDifference();
        }

        private void OpenDifference()
        {
            FileForm form = new FileForm();
            if (form.ShowDialog(this) == DialogResult.OK)
            {
                CompareExcel(form.leftPath, form.rightPath);
            }
        }

        private void UpdateNextPreviousButton()
        {
            nextButton.Enabled = GetNextDifferenctRowIndex() != -1;
            previousButton.Enabled = GetPreviousDifferenctRowIndex() != -1;
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                if (e.KeyCode == Keys.P)
                {
                    SelectPreviousDifference();
                }
                else if (e.KeyCode == Keys.N)
                {
                    SelectNextDifference();
                }
                else if (e.KeyCode == Keys.O)
                {
                    OpenDifference();
                }
            }
        }
    }
}
