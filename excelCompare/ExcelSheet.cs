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

        public string GetContent(int columnCount)
        {
            StringBuilder builder = new StringBuilder();
            for ( int i = 0; i < _rows.Count; i++ )
            {
                if (i > 0)
                {
                    builder.Append(SheetComparer.LINE_SPLITTER);
                }
                _rows[i].GetContent(columnCount, builder);
            }
            return builder.ToString();
        }

        public void SetRowTarget(int rowIndex, int targetIndex, bool changed)
        {
            _rows[rowIndex].targetIndex = targetIndex;
            _rows[rowIndex].different = changed;
            _rowMap[targetIndex] = rowIndex;
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

        public ExcelRow GetRowByRealIndex(int rowIndex)
        {
            if (rowIndex >= 0 && rowIndex < _rows.Count)
            {
                return _rows[rowIndex];
            }
            return null;
        }
    }
}
