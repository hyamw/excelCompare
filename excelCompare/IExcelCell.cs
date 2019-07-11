using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    /// <summary>
    /// 单元格
    /// </summary>
    internal interface IExcelCell
    {
        /// <summary>
        /// 该单元格所在的行序号
        /// </summary>
        int rowIndex
        {
            get;
        }

        /// <summary>
        /// 该单元格所在的列序号
        /// </summary>
        int columnIndex
        {
            get;
        }

        /// <summary>
        /// 该单元格的文字
        /// </summary>
        string value
        {
            get;
        }

        /// <summary>
        /// 单元格类型
        /// </summary>
        CellType cellType
        {
            get;
        }

        /// <summary>
        /// 返回用于比较的字符串
        /// </summary>
        /// <returns></returns>
        string GetContent();

        /// <summary>
        /// 将内容保存到Excel单元格对象中
        /// </summary>
        /// <param name="cell"></param>
        void Save(ICell cell);
    }
}
