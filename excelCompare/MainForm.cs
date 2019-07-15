using DiffMatchPatch;
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
    public partial class MainForm : Form, IDifferenceAdapter
    {
        private Color DIFF_BACK_COLOR = Color.FromArgb(255, 227, 227);
        private Color DIFF_FORE_COLOR = Color.FromArgb(255, 0, 0);
        private bool stopUpdate = false;
        private DataTable lineTable = new DataTable();
        private ExcelWorkbook leftWorkbook;
        private ExcelWorkbook rightWorkbook;
        private SheetComparer currentSheet = null;
        private int _leftAlignIndex = -1;
        private int _rightAlignIndex = -1;
        private DataGridView updatingGrid = null;
        private Dictionary<string, bool> _sheetDifferences = new Dictionary<string, bool>();
        private DataGridView _selectedGrid = null;

        public MainForm(string[] args)
        {
            InitializeComponent();
            compareToolStripButton.Enabled = false;

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
                    LoadExcel(args[0], args[1]);
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

            differenceScrollBar.DataSource = this;
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
                    int realIndex = ((VRow)currentSheet.left.GetRow(e.Row.Index)).realRowIndex;
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
                    int realIndex = ((VRow)currentSheet.right.GetRow(e.Row.Index)).realRowIndex;
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

            if ( source == leftGrid )
            {
                differenceScrollBar.Value = source.FirstDisplayedScrollingRowIndex;
            }
        }

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
                for (int i = 0; i < source.SelectedRows.Count; i++)
                {
                    target.Rows[source.SelectedRows[i].Index].Cells[0].Selected = true;
                }
                updatingGrid = null;

                ShowLineDifference(source.SelectedCells[0].RowIndex);
                UpdateNextPreviousButton();
            }
            UpdateGridViewFocus();
            if ( source.CurrentCell != null )
            {
                differenceScrollBar.Selection = source.CurrentCell.RowIndex;
            }
        }

        private void UpdateGridViewFocus()
        {
            if (_selectedGrid != null && _selectedGrid.SelectedCells.Count > 0)
            {
                if (_selectedGrid == leftGrid)
                {
                    copyToolstripButton.Image = global::excelCompare.Properties.Resources.right;
                    copyToolstripButton.Enabled = true;
                    copyToolstripButton.ToolTipText = "复制到右边(Ctrl+R)";
                }
                else
                {
                    copyToolstripButton.Image = global::excelCompare.Properties.Resources.left;
                    copyToolstripButton.Enabled = true;
                    copyToolstripButton.ToolTipText = "复制到左边(Ctrl+L)";
                }
            }
            else
            {
                copyToolstripButton.Enabled = false;
            }

            saveToolStripMenuItem.Enabled = _selectedGrid != null;
            saveAsToolStripMenuItem.Enabled = _selectedGrid != null;

            if ( _selectedGrid != null )
            {
                if ( _selectedGrid == leftGrid )
                {
                    saveToolStripMenuItem.Text = "保存(左)";
                    saveAsToolStripMenuItem.Text = "另存为(左)";
                }
                else
                {
                    saveToolStripMenuItem.Text = "保存(右)";
                    saveAsToolStripMenuItem.Text = "另存为(右)";
                }
            }

            alignToolStripMenuItem.Enabled = currentSheet != null;
            leftToolStripMenuItem.Enabled = _selectedGrid == rightGrid && currentSheet != null && rightGrid.SelectedRows.Count > 0;
            rightToolStripMenuItem.Enabled = _selectedGrid == leftGrid && currentSheet != null && leftGrid.SelectedRows.Count > 0;
            copy2LeftMenuItem.Enabled = leftToolStripMenuItem.Enabled;
            copy2RightMenuItem.Enabled = rightToolStripMenuItem.Enabled;

            alignMenuItem.Enabled = leftGrid.SelectedCells.Count > 0 || rightGrid.SelectedCells.Count > 0;
        }

        private void ClearAlignment()
        {
            leftGrid.Cursor = Cursors.Default;
            rightGrid.Cursor = Cursors.Default;
            _leftAlignIndex = -1;
            _rightAlignIndex = -1;
        }

        private void CompareSheet(string leftSheetName, string rightSheetName)
        {
            ClearAlignment();
            currentSheet = new SheetComparer();
            currentSheet.Execute(leftWorkbook.LoadSheet(leftSheetName), rightWorkbook.LoadSheet(rightSheetName));
            stopUpdate = true;

            leftGrid.DataSource = currentSheet.left.GetSource();
            rightGrid.DataSource = currentSheet.right.GetSource();
            stopUpdate = false;

            lineTable = new DataTable();

            for ( int i = 0; i < currentSheet.columnCount; i++ )
            {
                lineTable.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(i)));
            }

            rowDiffGrid.DataSource = lineTable;

            UpdateNextPreviousButton();
            AdjustRowHeaderSize();
            compareToolStripButton.Enabled = false;
            UpdateGridViewFocus();
            differenceScrollBar.ForceUpdate();
            differenceScrollBar.Selection = -1;
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
            differenceScrollBar.Selection = rowIndex;
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

        private string SafeGetColumn(IExcelRow row, int columnIndex)
        {
            if ( row == null )
            {
                return string.Empty;
            }
            IExcelCell cell = row.GetCell(columnIndex);
            if (cell != null && cell.value != null)
            {
                return cell.value;
            }
            return string.Empty;
        }

        private void ShowLineDifference(int rowIndex)
        {
            rowDiffGrid.DataSource = null;

            IExcelRow leftRow = currentSheet.left.GetRow(rowIndex);
            IExcelRow rightRow = currentSheet.right.GetRow(rowIndex);
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

        private void GetSummary()
        {
            _sheetDifferences.Clear();
            for (int i = 0; i < leftWorkbook.sheetNames.Count; i++)
            {
                string sheetName = leftWorkbook.sheetNames[i];
                if (rightWorkbook.sheetNames.Contains(sheetName))
                {
                    ExcelSheet leftExcelSheet = leftWorkbook.LoadSheet(sheetName);
                    ExcelSheet rightExcelSheet = rightWorkbook.LoadSheet(sheetName);
                    int columnCount = Math.Max(leftExcelSheet.columnCount, rightExcelSheet.columnCount);
                    VSheet leftSheet = new VSheet(leftExcelSheet, columnCount);
                    VSheet rightSheet = new VSheet(rightExcelSheet, columnCount);

                    diff_match_patch comparer = new diff_match_patch();
                    string leftContent = leftSheet.GetContent();
                    string rightContent = rightSheet.GetContent();
                    List<Diff> diffs = comparer.diff_main(leftContent, rightContent, true);
                    comparer.diff_cleanupSemanticLossless(diffs);

                    bool isDifferent = false;
                    for ( int diffIndex = 0; diffIndex < diffs.Count; diffIndex++ )
                    {
                        if (diffs[diffIndex].operation != Operation.EQUAL)
                        {
                            isDifferent = true;
                            break;
                        }
                    }

                    _sheetDifferences.Add(sheetName, isDifferent);
                }
            }
        }

        private bool IsSheetDifferent(string name)
        {
            bool different = false;
            if ( !_sheetDifferences.TryGetValue(name, out different) )
            {
                different = false;
            }
            return different;
        }

        private void LoadExcel(string leftPath, string rightPath)
        {
            leftWorkbook = new ExcelWorkbook();
            leftWorkbook.Load(leftPath);

            rightWorkbook = new ExcelWorkbook();
            rightWorkbook.Load(rightPath);

            GetSummary();

            CompareSheet(leftWorkbook.sheetNames[0], rightWorkbook.sheetNames[0]);

            leftComboBox.Items.Clear();

            for (int i = 0; i < leftWorkbook.sheetNames.Count; i++)
            {
                string sheetName = leftWorkbook.sheetNames[i];
                leftComboBox.Items.Add(sheetName + (IsSheetDifferent(sheetName) ? " *" : string.Empty));
            }

            leftComboBox.SelectedIndex = 0;

            rightComboBox.Items.Clear();

            for (int i = 0; i < rightWorkbook.sheetNames.Count; i++)
            {
                string sheetName = rightWorkbook.sheetNames[i];
                rightComboBox.Items.Add(sheetName + (IsSheetDifferent(sheetName) ? " *" : string.Empty));
            }

            rightComboBox.SelectedIndex = 0;

            leftFileTextBox.Text = leftPath;
            rightFileTextBox.Text = rightPath;
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
                LoadExcel(form.leftPath, form.rightPath);
            }
        }

        private void UpdateNextPreviousButton()
        {
            nextButton.Enabled = GetNextDifferenctRowIndex() != -1;
            previousButton.Enabled = GetPreviousDifferenctRowIndex() != -1;
            nextToolStripMenuItem.Enabled = GetNextDifferenctRowIndex() != -1;
            previousToolStripMenuItem.Enabled = GetPreviousDifferenctRowIndex() != -1;
        }

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                ClearAlignment();
            }
        }

        private void PrepareAlignment()
        {
            if (leftGrid.Focused && leftGrid.SelectedCells.Count > 0)
            {
                _leftAlignIndex = leftGrid.CurrentCell.RowIndex;
                leftGrid.Cursor = Cursors.No;
                rightGrid.Cursor = Cursors.Help;
            }
            else if (rightGrid.Focused && rightGrid.SelectedCells.Count > 0)
            {
                _rightAlignIndex = rightGrid.CurrentCell.RowIndex;
                leftGrid.Cursor = Cursors.Help;
                rightGrid.Cursor = Cursors.No;
            }
        }

        private void OnAlignMenuClicked(object sender, EventArgs e)
        {
            PrepareAlignment();
        }

        private void CompareWithAlignment()
        {
            int leftRowIndex = leftGrid.FirstDisplayedScrollingRowIndex;
            leftGrid.Cursor = Cursors.Default;
            rightGrid.Cursor = Cursors.Default;
            SheetComparer sheetComparer = new SheetComparer();
            sheetComparer.Execute(currentSheet.left, currentSheet.right, _leftAlignIndex, _rightAlignIndex);
            stopUpdate = true;

            currentSheet = sheetComparer;

            int visibleRow = Math.Min(_leftAlignIndex, _rightAlignIndex);

            _leftAlignIndex = -1;
            _rightAlignIndex = -1;

            leftGrid.DataSource = currentSheet.left.GetSource();
            rightGrid.DataSource = currentSheet.right.GetSource();
            stopUpdate = false;

            lineTable = new DataTable();

            for (int i = 0; i < currentSheet.columnCount; i++)
            {
                lineTable.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(i)));
            }

            rowDiffGrid.DataSource = lineTable;

            if (visibleRow != -1)
            {
                JumToRow(visibleRow);
            }
            else
            {
                UpdateNextPreviousButton();
            }

            leftGrid.FirstDisplayedScrollingRowIndex = leftRowIndex;
            UpdateGridViewFocus();
            differenceScrollBar.ForceUpdate();
        }

        private void OnCellClicked(object sender, DataGridViewCellEventArgs e)
        {
            if ( sender == rightGrid && _leftAlignIndex != -1 )
            {
                _rightAlignIndex = e.RowIndex;

                CompareWithAlignment();
            }
            else if (sender == leftGrid && _rightAlignIndex != -1)
            {
                _leftAlignIndex = e.RowIndex;

                CompareWithAlignment();
            }
        }

        private void OnCellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            DataGridView source = GetSource(sender);
            DataGridView target = GetTarget(sender);
            if ( e.Button == MouseButtons.Right )
            {
                bool multiSelect = source.SelectedRows.Count > 1 || (Control.ModifierKeys & Keys.Control) == Keys.Control;
                List<int> selectedRows = new List<int>();
                if (multiSelect)
                {
                    for (int i = 0; i < source.SelectedRows.Count; i++)
                    {
                        selectedRows.Add(source.SelectedRows[i].Index);
                    }
                }

                selectedRows.Add(e.RowIndex);
                stopUpdate = true;
                source.CurrentCell = source.Rows[e.RowIndex].Cells[0];
                for ( int i = 0; i < selectedRows.Count; i++ )
                {
                    source.Rows[selectedRows[i]].Cells[0].Selected = true;
                    target.Rows[selectedRows[i]].Cells[0].Selected = true;
                }
                stopUpdate = false;
            }
            _selectedGrid = source;
            UpdateGridViewFocus();
        }

        private void OnCellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if ( (e.State & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected)
            {
                if (currentSheet.IsCellDifferent(e.RowIndex, e.ColumnIndex))
                {
                    e.CellStyle.SelectionForeColor = DIFF_FORE_COLOR;
                }
                else
                {
                    e.CellStyle.SelectionForeColor = Color.Black;
                }
            }
        }

        private void OnRowGridCellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if ((e.State & DataGridViewElementStates.Selected) == DataGridViewElementStates.Selected)
            {
                if (CompareColumn(e.ColumnIndex))
                {
                    e.CellStyle.SelectionForeColor = DIFF_FORE_COLOR;
                }
                else
                {
                    e.CellStyle.SelectionForeColor = Color.Black;

                }
            }
        }

        private void AdjustRowHeaderSize()
        {
            int maxRowIndex = 0;
            if (currentSheet != null)
            {
                maxRowIndex = currentSheet.rowCount;
            }
            Graphics graphics = leftGrid.CreateGraphics();
            int width = (int)graphics.MeasureString(maxRowIndex.ToString(), leftGrid.Font).Width;
            leftGrid.RowHeadersWidth = width + 35;
            rightGrid.RowHeadersWidth = width + 35;
        }

        private void OnCompareButtonClicked(object sender, EventArgs e)
        {
            int leftSelectionIndex = leftComboBox.SelectedIndex;
            int rightSelectionIndex = rightComboBox.SelectedIndex;
            if (leftSelectionIndex == -1 || rightSelectionIndex == -1)
            {
                return;
            }

            CompareSheet(leftWorkbook.sheetNames[leftSelectionIndex], rightWorkbook.sheetNames[rightSelectionIndex]);
        }

        private string GetLeftSheetName()
        {
            string sheetName = string.Empty;
            if (leftWorkbook != null && leftComboBox.SelectedIndex >= 0 && leftComboBox.SelectedIndex < leftWorkbook.sheetNames.Count)
            {
                sheetName = leftWorkbook.sheetNames[leftComboBox.SelectedIndex];
            }
            return sheetName;
        }

        private string GetRightSheetName()
        {
            string sheetName = string.Empty;
            if (rightWorkbook != null && rightComboBox.SelectedIndex >= 0 && rightComboBox.SelectedIndex < rightWorkbook.sheetNames.Count)
            {
                sheetName = rightWorkbook.sheetNames[rightComboBox.SelectedIndex];
            }
            return sheetName;
        }

        private void OnSheetSelectionChanged(object sender, EventArgs e)
        {
            if (currentSheet == null)
            {
                return;
            }
            string leftName = GetLeftSheetName();
            string rightName = GetRightSheetName();
            ComboBox target = null;
            if ( sender == leftComboBox )
            {
                target = rightComboBox;
            }
            else
            {
                target = leftComboBox;
            }
            compareToolStripButton.Enabled = leftName != currentSheet.left.name || rightName != currentSheet.right.name;

            if (sender == leftComboBox && _sheetDifferences.ContainsKey(leftName))
            {
                target.SelectedItem = ((ComboBox)sender).SelectedItem;
            }
        }

        private void Copy2Left()
        {
            if ( rightGrid.SelectedRows.Count > 0 )
            {
                int firstRowIndex = rightGrid.FirstDisplayedScrollingRowIndex;
                DataTable table = leftGrid.DataSource as DataTable;
                List<int> selectionList = new List<int>();
                for ( int selectionIndex = 0; selectionIndex < rightGrid.SelectedRows.Count; selectionIndex++ )
                {
                    int rowIndex = rightGrid.SelectedRows[selectionIndex].Index;
                    currentSheet.Copy2Left(rowIndex);
                    VRow row = currentSheet.left.GetRow(rowIndex) as VRow;

                    DataRow dataRow = table.Rows[rowIndex];
                    dataRow.BeginEdit();
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        IExcelCell cell = row.GetCell(i);
                        dataRow[i] = cell != null ? cell.value : string.Empty;
                    }
                    dataRow.EndEdit();
                    selectionList.Add(rowIndex);
                }
                table.AcceptChanges();

                rightGrid.FirstDisplayedScrollingRowIndex = firstRowIndex;

                stopUpdate = true;
                leftGrid.ClearSelection();
                for ( int i = 0; i < selectionList.Count; i++ )
                {
                    rightGrid.Rows[selectionList[i]].Cells[0].Selected = true;
                    leftGrid.Rows[selectionList[i]].Cells[0].Selected = true;
                }
                stopUpdate = false;
            }
        }

        private void Copy2Right()
        {
            if (leftGrid.SelectedCells.Count > 0)
            {
                int firstRowIndex = leftGrid.FirstDisplayedScrollingRowIndex;
                DataTable table = rightGrid.DataSource as DataTable;
                List<int> selectionList = new List<int>();
                for (int selectionIndex = 0; selectionIndex < leftGrid.SelectedRows.Count; selectionIndex++)
                {
                    int rowIndex = leftGrid.SelectedRows[selectionIndex].Index;
                    currentSheet.Copy2Right(rowIndex);
                    VRow row = currentSheet.right.GetRow(rowIndex) as VRow;

                    DataRow dataRow = table.Rows[rowIndex];
                    dataRow.BeginEdit();
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        IExcelCell cell = row.GetCell(i);
                        dataRow[i] = cell != null ? cell.value : string.Empty;
                    }
                    dataRow.EndEdit();
                    selectionList.Add(rowIndex);
                }
                table.AcceptChanges();
                leftGrid.FirstDisplayedScrollingRowIndex = firstRowIndex;

                stopUpdate = true;
                rightGrid.ClearSelection();
                for (int i = 0; i < selectionList.Count; i++)
                {
                    leftGrid.Rows[selectionList[i]].Cells[0].Selected = true;
                    rightGrid.Rows[selectionList[i]].Cells[0].Selected = true;
                }
                stopUpdate = false;
            }
        }

        private void Save(ExcelWorkbook workbook, VSheet sheet, bool saveAs)
        {
            string outputPath = null;
            if (saveAs)
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = FileForm.FILE_FILTER;
                dialog.Title = "另存为";
                dialog.FileName = Path.GetFileName(workbook.filePath);
                dialog.InitialDirectory = Path.GetDirectoryName(workbook.filePath);
                dialog.FilterIndex = FileForm.GetFilterIndex(workbook.filePath);
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    outputPath = dialog.FileName;
                }
                else
                {
                    return;
                }
            }

            try
            {
                workbook.SaveSheet(sheet, outputPath);
            }
            catch(Exception ex)
            {
                MessageBox.Show(this, ex.ToString());
            }
        }

        private void SaveLeft(bool saveAs = false)
        {
            Save(leftWorkbook, currentSheet.left, saveAs);
        }

        private void SaveRight(bool saveAs = false)
        {
            Save(rightWorkbook, currentSheet.right, saveAs);
        }

        private void OnGridViewFocusEnter(object sender, EventArgs e)
        {
            _selectedGrid = GetSource(sender);
            UpdateGridViewFocus();
        }

        private void OnGridViewFocusLeave(object sender, EventArgs e)
        {
            UpdateGridViewFocus();
        }

        private void OnCopy2LeftClicked(object sender, EventArgs e)
        {
            Copy2Left();
        }

        private void OnCopy2RightClicked(object sender, EventArgs e)
        {
            Copy2Right();
        }

        private void OnCopyButtonClicked(object sender, EventArgs e)
        {
            if ( _selectedGrid == leftGrid )
            {
                Copy2Right();
            }
            else if ( _selectedGrid == rightGrid )
            {
                Copy2Left();
            }
        }

        private void OnSaveMenuClicked(object sender, EventArgs e)
        {
            if ( _selectedGrid == null )
            {
                return;
            }
            if ( _selectedGrid == leftGrid )
            {
                SaveLeft(false);
            }
            else
            {
                SaveRight(false);
            }
        }

        private void OnSaveAsMenuClicked(object sender, EventArgs e)
        {
            if (_selectedGrid == null)
            {
                return;
            }
            if (_selectedGrid == leftGrid)
            {
                SaveLeft(true);
            }
            else
            {
                SaveRight(true);
            }
        }

        private void OnExitMenuClicked(object sender, EventArgs e)
        {
            this.Close();
        }

        public int DifferenceCount
        {
            get
            {
                if ( currentSheet != null )
                {
                    return currentSheet.rowCount;
                }
                return 0;
            }
        }

        public float PageSizeRatio
        {
            get
            {
                if ( leftGrid.Rows.Count > 0 )
                {
                    return (float)GetVisibleRowCount() / (float)leftGrid.Rows.Count;
                }
                return (float)leftGrid.Size.Height / (float)leftGrid.PreferredSize.Height;
            }
        }

        public DifferentMode GetDifference(int index)
        {
            return currentSheet.GetRowDifference(index);
        }

        private void OnDifferenceScrolled(object sender, ScrollEventArgs e)
        {
            leftGrid.FirstDisplayedScrollingRowIndex = Math.Min(leftGrid.Rows.Count - 1, e.NewValue);
        }

        private int GetVisibleRowCount()
        {
            int visibleRowCount = 0;
            if (leftGrid.GetHorizontalScrollBar().Visible)
            {
                visibleRowCount = (int)Math.Floor((leftGrid.Size.Height - leftGrid.GetHorizontalScrollBar().Size.Height + 2) / (float)leftGrid.Rows[0].Cells[0].Size.Height) - 1;
            }
            else
            {
                visibleRowCount = (int)Math.Floor(leftGrid.Size.Height / (float)leftGrid.Rows[0].Cells[0].Size.Height) - 1;
            }
            if (visibleRowCount > leftGrid.Rows.Count)
            {
                visibleRowCount = leftGrid.Rows.Count;
            }
            return visibleRowCount;
        }
    }
}
