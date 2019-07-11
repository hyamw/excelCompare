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
    internal class ExcelSheet : IExcelSheet
    {
        private string _name = string.Empty;
        private List<ExcelRow> _rows = new List<ExcelRow>();
        private int _columnCount = 0;

        public string name
        {
            get
            {
                return _name;
            }
        }

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

        public int rowCount
        {
            get
            {
                return _rows.Count;
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

        private string GetCachedCellValue(ICell cell)
        {
            switch( cell.CachedFormulaResultType)
            {
                case CellType.Blank:
                    return string.Empty;
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.String:
                    return cell.StringCellValue;
                default:
                    return null;
            }
        }

        private string GetCellValue(ICell cell, IFormulaEvaluator evaluator)
        {
            if ( cell != null )
            {
                if ( cell.CellType == CellType.Formula )
                {
                    try
                    {
                        return evaluator.Evaluate(cell).FormatAsString();
                    }
                    catch(Exception)
                    {
                        return GetCachedCellValue(cell);
                    }
                }
                else
                {
                    return cell.ToString();
                }
            }
            return null;
        }

        private CellType GetCellType(ICell cell)
        {
            if (cell != null)
            {
                return cell.CellType;
            }
            return CellType.Blank;
        }

        public bool Load(ISheet sheet)
        {
            IFormulaEvaluator evaluator = sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
            _name = sheet.SheetName;
            bool success = false;
            _columnCount = GetSheetColumnCount(sheet);

            for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                ExcelRow newRow = new ExcelRow(this, rowIndex);
                if (row != null)
                {
                    for (int k = 0; k < columnCount; k++)
                    {
                        ICell cell = row.GetCell(k);
                        newRow.AddCell(new ExcelCell(rowIndex, k, GetCellType(cell), GetCellValue(cell, evaluator)));
                    }
                }
                else
                {
                    for (int k = 0; k < columnCount; k++)
                    {
                        newRow.AddCell(new ExcelCell(rowIndex, k, CellType.Blank, null));
                    }
                }
                _rows.Add(newRow);
            }
            return success;
        }

        public ExcelRow GetRowByRealIndex(int rowIndex)
        {
            if (rowIndex >= 0 && rowIndex < _rows.Count)
            {
                return _rows[rowIndex];
            }
            return null;
        }

        public string GetContent()
        {
            throw new NotImplementedException();
        }

        public string GetContent(int beginRow, int endRow)
        {
            throw new NotImplementedException();
        }

        public IExcelRow GetRow(int rowIndex)
        {
            if ( rowIndex >= 0 && rowIndex < _rows.Count )
            {
                return _rows[rowIndex];
            }
            return null;
        }
    }
}
