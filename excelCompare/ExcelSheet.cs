using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    public class ExcelSheet
    {
        private List<ExcelRow> _rows = new List<ExcelRow>();
        private int _columnCount = 0;
        private int _diffViewRowCount = 0;
        private int _diffViewColumnCount = 0;
        private Dictionary<int, int> _rowMap = new Dictionary<int, int>();

        public List<ExcelRow> rows
        {
            get
            {
                return _rows;
            }
        }

        public int columnCount
        {
            get
            {
                return _columnCount;
            }
        }

        public int diffViewRowCount
        {
            get
            {
                return _diffViewRowCount;
            }
        }

        public int diffViewColumnCount
        {
            get
            {
                return _diffViewColumnCount;
            }
        }

        private int GetSheetColumnCount(ISheet sheet)
        {
            int cellNum = 0;
            for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }
                int lastColumn = sheet.GetRow(rowIndex).LastCellNum;
                if (lastColumn > cellNum)
                {
                    cellNum = lastColumn;
                }
            }
            return cellNum;
        }

        //public bool LoadExcel(string filePath, int? sheetIndex = null)
        //{
        //    bool success = false;
        //    IWorkbook fileWorkbook;
        //    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        //    {
        //        if (filePath.Last() == 's')
        //        {
        //            try
        //            {
        //                fileWorkbook = new HSSFWorkbook(fs);
        //            }
        //            catch (Exception ex)
        //            {
        //                throw ex;
        //            }
        //        }
        //        else
        //        {
        //            try
        //            {
        //                fileWorkbook = new XSSFWorkbook(fs);
        //            }
        //            catch (Exception ex)
        //            {
        //                fileWorkbook = new HSSFWorkbook(fs);
        //            }
        //        }
        //    }

        //    for (int i = 0; i < fileWorkbook.NumberOfSheets; i++)
        //    {
        //        if (sheetIndex != null && sheetIndex != i)
        //            continue;
        //        ISheet sheet = fileWorkbook.GetSheetAt(i);

        //        _columnCount = GetSheetColumnCount(sheet);

        //        for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
        //        {
        //            IRow row = sheet.GetRow(rowIndex);
        //            ExcelRow newRow = new ExcelRow(rowIndex);
        //            if (row != null)
        //            {
        //                for (int k = 0; k < columnCount; k++)
        //                {
        //                    ICell cell = row.GetCell(k);
        //                    newRow.columns.Add(cell != null ? cell.ToString() : "");
        //                }
        //            }
        //            else
        //            {
        //                for (int k = 0; k < columnCount; k++)
        //                {
        //                    newRow.columns.Add("");
        //                }
        //            }
        //            newRow.Compile();
        //            _rows.Add(newRow);
        //        }
        //        success = true;
        //        _diffViewRowCount = _rows.Count;
        //    }
        //    return success;
        //}

        public bool Load(ISheet sheet)
        {
            bool success = false;
            _columnCount = GetSheetColumnCount(sheet);

            for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                ExcelRow newRow = new ExcelRow(rowIndex);
                if (row != null)
                {
                    for (int k = 0; k < columnCount; k++)
                    {
                        ICell cell = row.GetCell(k);
                        newRow.columns.Add(cell != null ? cell.ToString() : "");
                    }
                }
                else
                {
                    for (int k = 0; k < columnCount; k++)
                    {
                        newRow.columns.Add("");
                    }
                }
                newRow.Compile();
                _rows.Add(newRow);
            }
            _diffViewRowCount = _rows.Count;
            return success;
        }

        public string GetContent()
        {
            StringBuilder builder = new StringBuilder();
            for ( int i = 0; i < _rows.Count; i++ )
            {
                if (i > 0)
                {
                    builder.Append(DiffAnalyzer.LINE_SPLITTER);
                }
                builder.AppendLine(_rows[i].signature);
            }
            return builder.ToString();
        }

        public void SetRowTarget(int rowIndex, int targetIndex, bool changed)
        {
            _rows[rowIndex].targetIndex = targetIndex;
            _rows[rowIndex].different = changed;
            _rowMap[targetIndex] = rowIndex;
        }

        public bool IsRowDifferent(int targetIndex)
        {
            int rowIndex = -1;
            if ( _rowMap.TryGetValue(targetIndex, out rowIndex) )
            {
                return _rows[rowIndex].different;
            }
            return true;
        }

        public bool CompareCell(ExcelSheet other, int rowIndex, int columnIndex)
        {
            if ( IsRowDifferent(rowIndex) || other.IsRowDifferent(rowIndex) )
            {
                ExcelRow selfRow = GetRow(rowIndex);
                ExcelRow otherRow = other.GetRow(rowIndex);
                if (selfRow == null || otherRow == null)
                {
                    return true;
                }

                string selfCell = selfRow.GetColumn(columnIndex);
                string otherCell = otherRow.GetColumn(columnIndex);
                if ((selfCell == null && otherCell != null) || (selfCell != null && otherCell == null))
                {
                    return true;
                }

                return string.Compare(selfCell, otherCell) != 0;
            }
            return false;
        }

        private int GetRealIndex(int index)
        {
            int rowIndex = -1;
            if (!_rowMap.TryGetValue(index, out rowIndex))
            {
                rowIndex = -1;
            }
            return rowIndex;
        }

        public ExcelRow GetRow(int rowIndex)
        {
            int realIndex = GetRealIndex(rowIndex);
            if ( realIndex == -1 )
            {
                return null;
            }
            return _rows[realIndex];
        }

        public void SetDiffViewSize(int rowCount, int columnCount)
        {
            _diffViewRowCount = rowCount;
            _diffViewColumnCount = columnCount;
        }

        public DataTable GenerateSource()
        {
            DataTable dt = new DataTable();
            for (int j = 0; j < _diffViewColumnCount; j++)
            {
                dt.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(j)));
            }

            int rowIndex = 0;
            DataRow dr = null;
            for ( int i = 0; i < _diffViewRowCount; i++ )
            {
                dr = dt.NewRow();
                if (_rows.Count > rowIndex && i == _rows[rowIndex].targetIndex )
                {
                    for (int k = 0; k < _diffViewColumnCount; k++)
                    {
                        if (k < columnCount)
                        {
                            dr[k] = _rows[rowIndex].columns[k];
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

        public void PaddColumns(int columnCount)
        {
            _columnCount = columnCount;
            for ( int i = 0; i < _rows.Count; i++ )
            {
                _rows[i].PaddColumns(columnCount);
            }
        }
    }
}
