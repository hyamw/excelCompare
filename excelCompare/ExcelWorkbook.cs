using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class ExcelWorkbook
    {
        private IWorkbook workbook = null;
        private List<string> _sheetNames = new List<string>();

        public List<string> sheetNames
        {
            get
            {
                return _sheetNames;
            }
        }

        public void Load(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (filePath.Last() == 's')
                {
                    try
                    {
                        workbook = new HSSFWorkbook(fs);
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
                        workbook = new XSSFWorkbook(fs);
                    }
                    catch (Exception)
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                }
            }

            if (workbook != null)
            {
                for ( int i = 0; i < workbook.NumberOfSheets; i++ )
                {
                    _sheetNames.Add(workbook.GetSheetName(i));
                }
            }
        }

        public ExcelSheet LoadSheet(string name)
        {
            return LoadSheet(workbook.GetSheetIndex(name));
        }

        public ExcelSheet LoadSheet(int sheetIndex)
        {
            if (sheetIndex >= 0 && sheetIndex < workbook.NumberOfSheets)
            {
                ISheet sheet = workbook.GetSheetAt(sheetIndex);
                ExcelSheet sheetWrapper = new ExcelSheet();
                sheetWrapper.Load(sheet);
                return sheetWrapper;
            }
            return new ExcelSheet();
        }
    }
}
