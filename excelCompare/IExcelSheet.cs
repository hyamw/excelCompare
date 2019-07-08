using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal interface IExcelSheet
    {
        /// <summary>
        /// 返回当前表格的名字
        /// </summary>
        string name
        {
            get;
        }

        /// <summary>
        /// 返回当前表格的列数
        /// </summary>
        int columnCount
        {
            get;
        }

        /// <summary>
        /// 返回当前表格的行数
        /// </summary>
        int rowCount
        {
            get;
        }

        /// <summary>
        /// 生成比较用的内容
        /// </summary>
        /// <returns></returns>
        string GetContent();

        /// <summary>
        /// 将指定范围的行生成比较用的内容
        /// </summary>
        /// <param name="beginRow">起始行(包含该行)</param>
        /// <param name="endRow">结束行(包含该行)</param>
        /// <returns></returns>
        string GetContent(int beginRow, int endRow);

        /// <summary>
        /// 获取指定行
        /// </summary>
        /// <param name="rowIndex">指定行序号</param>
        /// <returns>如果行存在则返回对应行的对象, 否则返回null</returns>
        IExcelRow GetRow(int rowIndex);
    }
}
