using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class VSheet : IExcelSheet
    {
        private string _name = string.Empty;
        private List<IExcelRow> _rows = new List<IExcelRow>();
        private int _columnCount = 0;

        /// <summary>
        /// 返回当前表格的名字
        /// </summary>
        public string name
        {
            get
            {
                return _name;
            }
        }

        /// <summary>
        /// 返回当前表格的列数
        /// </summary>
        public int columnCount
        {
            get
            {
                return _columnCount;
            }
        }

        /// <summary>
        /// 返回当前表格的行数
        /// </summary>
        public int rowCount
        {
            get
            {
                return _rows.Count;
            }
        }

        /// <summary>
        /// 生成比较用的内容
        /// </summary>
        /// <returns></returns>
        public string GetContent()
        {
            return GetContent(0, rowCount - 1);
        }

        /// <summary>
        /// 将指定范围的行生成比较用的内容
        /// </summary>
        /// <param name="beginRow">起始行(包含该行)</param>
        /// <param name="endRow">结束行(包含该行)</param>
        /// <returns></returns>
        public string GetContent(int beginRow, int endRow)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = beginRow; i <= endRow; i++)
            {
                _rows[i].BuildContent(builder);
                builder.AppendLine(SheetComparer.LINE_SPLITTER);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 获取指定行
        /// </summary>
        /// <param name="rowIndex">指定行序号</param>
        /// <returns>如果行存在则返回对应行的对象, 否则返回null</returns>
        public IExcelRow GetRow(int rowIndex)
        {
            if (rowIndex >= 0 && rowIndex < _rows.Count)
            {
                return _rows[rowIndex];
            }
            return null;
        }

        /// <summary>
        /// 根据IExcelSheet构造一个用于对比的VSheet对象
        /// </summary>
        /// <param name="sheet"></param>
        public VSheet(IExcelSheet sheet, int columnCount)
        {
            _name = sheet.name;
            _columnCount = columnCount;
            for ( int i = 0; i < sheet.rowCount; i++ )
            {
                _rows.Add(new VRow(this, i, sheet.GetRow(i)));
            }
        }

        internal VSheet(string sheetName, int columnCount)
        {
            _name = sheetName;
            _columnCount = columnCount;
        }

        /// <summary>
        /// 从指定的行序号开始填充rowCount个空行
        /// </summary>
        /// <param name="startRowIndex">开始填充的行序号</param>
        /// <param name="rowCount">需要填充的空行数量</param>
        public void PadRowsAt(int startRowIndex, int rowCount)
        {
            for ( int i = 0; i < rowCount; i++ )
            {
                _rows.Insert(startRowIndex, new VRow(this, i, null));
            }
        }

        /// <summary>
        /// 增加一个新的行
        /// </summary>
        /// <returns></returns>
        internal VRow NewRow()
        {
            VRow row = new VRow(this, _rows.Count, null);
            _rows.Add(row);
            return row;
        }

        /// <summary>
        /// 增加一个新的行
        /// </summary>
        /// <returns></returns>
        internal VRow NewRow(IExcelRow referenceRow)
        {
            VRow row = new VRow(this, _rows.Count, referenceRow);
            _rows.Add(row);
            return row;
        }

        public DataTable GetSource()
        {
            DataTable dt = new DataTable();
            for (int j = 0; j < _columnCount; j++)
            {
                dt.Columns.Add(new DataColumn(CellReference.ConvertNumToColString(j)));
            }
            
            DataRow dr = null;
            for (int i = 0; i < rowCount; i++)
            {
                dr = dt.NewRow();
                IExcelRow row = _rows[i];
                for (int k = 0; k < _columnCount; k++)
                {
                    if (k < columnCount && row != null)
                    {
                        string displayText = row.GetCell(k).value;
                        if (displayText == null)
                        {
                            displayText = string.Empty;
                        }
                        dr[k] = displayText;
                    }
                    else
                    {
                        dr[k] = string.Empty;
                    }
                }
                
                dt.Rows.Add(dr);
            }
            return dt;
        }
    }
}
