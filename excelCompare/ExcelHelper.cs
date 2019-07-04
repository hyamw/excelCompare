using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace excelCompare
{
    public class ExcelHelper
    {
        /// <summary>
        /// 根据Excel和Sheet返回DataTable
        /// </summary>
        /// <param name="filePath">Excel文件地址</param>
        /// <param name="sheetIndex">Sheet索引</param>
        /// <returns>DataTable</returns>
        public static DataTable GetDataTable(string filePath, int sheetIndex)
        {
            return GetDataSet(filePath, sheetIndex).Tables[0];
        }

        public static int GetSheetColumnCount(ISheet sheet)
        {
            int cellNum = 0;
            for ( int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++ )
            {
                IRow row = sheet.GetRow(rowIndex);
                if ( row == null )
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

        /// <summary>
        /// 根据Excel返回DataSet
        /// </summary>
        /// <param name="filePath">Excel文件地址</param>
        /// <param name="sheetIndex">Sheet索引，可选，默认返回所有Sheet</param>
        /// <returns>DataSet</returns>
        public static DataSet GetDataSet(string filePath, int? sheetIndex = null)
        {
            DataSet ds = new DataSet();
            IWorkbook fileWorkbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (filePath.Last() == 's')
                {
                    try
                    {
                        fileWorkbook = new HSSFWorkbook(fs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                else
                {
                    try
                    {
                        fileWorkbook = new XSSFWorkbook(fs);
                    }
                    catch(Exception ex)
                    {
                        fileWorkbook = new HSSFWorkbook(fs);
                    }
                }
            }

            for (int i = 0; i < fileWorkbook.NumberOfSheets; i++)
            {
                if (sheetIndex != null && sheetIndex != i)
                    continue;
                DataTable dt = new DataTable();
                ISheet sheet = fileWorkbook.GetSheetAt(i);

                int columnCount = GetSheetColumnCount(sheet);
                for ( int j = 0; j < columnCount; j++ )
                {
                    dt.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(j)));
                }

                //数据
                for (int rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    DataRow dr = dt.NewRow();
                    if (row != null)
                    {
                        for (int k = 0; k < columnCount; k++)
                        {
                            ICell cell = row.GetCell(k);
                            dr[k] = cell != null ? cell.ToString() : "";
                        }
                    }
                    dt.Rows.Add(dr);
                }
                ds.Tables.Add(dt);
            }

            return ds;
        }

        public static List<ExcelRow> LoadExcel(string filePath, int? sheetIndex = null)
        {
            List<ExcelRow> list = new List<ExcelRow>();
            IWorkbook fileWorkbook;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (filePath.Last() == 's')
                {
                    try
                    {
                        fileWorkbook = new HSSFWorkbook(fs);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
                else
                {
                    try
                    {
                        fileWorkbook = new XSSFWorkbook(fs);
                    }
                    catch (Exception ex)
                    {
                        fileWorkbook = new HSSFWorkbook(fs);
                    }
                }
            }

            for (int i = 0; i < fileWorkbook.NumberOfSheets; i++)
            {
                if (sheetIndex != null && sheetIndex != i)
                    continue;
                ISheet sheet = fileWorkbook.GetSheetAt(i);

                int columnCount = GetSheetColumnCount(sheet);
                //for (int j = 0; j < columnCount; j++)
                //{
                //    dt.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(j)));
                //}

                //数据
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
                    list.Add(newRow);
                }
            }

            return list;
        }

        /// <summary>
        /// 根据DataTable导出Excel
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="file">保存地址</param>
        public static void GetExcelByDataTable(DataTable dt, string file)
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            GetExcelByDataSet(ds, file);
        }

        /// <summary>
        /// 根据DataSet导出Excel
        /// </summary>
        /// <param name="ds">DataSet</param>
        /// <param name="file">保存地址</param>
        public static void GetExcelByDataSet(DataSet ds, string file)
        {
            IWorkbook fileWorkbook = new HSSFWorkbook();
            int index = 0;
            foreach (DataTable dt in ds.Tables)
            {
                index++;
                ISheet sheet = fileWorkbook.CreateSheet("Sheet" + index);

                //表头
                IRow row = sheet.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ICell cell = row.CreateCell(i);
                    cell.SetCellValue(dt.Columns[i].ColumnName);
                }

                //数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    IRow row1 = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ICell cell = row1.CreateCell(j);
                        cell.SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            fileWorkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }

        /// <summary>
        /// 根据单元格将内容返回为对应类型的数据
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>数据</returns>
        private static object GetValueTypeForXLS(HSSFCell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank: //BLANK:
                    return null;
                case CellType.Boolean: //BOOLEAN:
                    return cell.BooleanCellValue;
                case CellType.Numeric: //NUMERIC:
                    return cell.NumericCellValue;
                case CellType.String: //STRING:
                    return cell.StringCellValue;
                case CellType.Error: //ERROR:
                    return cell.ErrorCellValue;
                case CellType.Formula: //FORMULA:
                default:
                    return "=" + cell.CellFormula;
            }
        }
    }
}
