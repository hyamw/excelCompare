using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    /// <summary>
    /// 该接口代表Excel中的一行
    /// </summary>
    internal interface IExcelRow
    {
        /// <summary>
        /// 该对象所属的表格对象
        /// </summary>
        IExcelSheet sheet
        {
            get;
        }

        /// <summary>
        /// 该行在表格中的行序号(从0开始)
        /// </summary>
        int rowIndex
        {
            get;
        }

        /// <summary>
        /// 获取指定列的单元格
        /// </summary>
        /// <param name="columnIndex">指定列序号(从0开始)</param>
        /// <returns>该列对应的单元格对象</returns>
        IExcelCell GetCell(int columnIndex);

        /// <summary>
        /// 生成比较用的内容
        /// </summary>
        /// <param name="builder">输出用的字符串构造器</param>
        void BuildContent(StringBuilder builder);
    }
}
