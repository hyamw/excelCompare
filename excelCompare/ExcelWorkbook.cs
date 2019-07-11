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
        private string _filePath = string.Empty;
        private IWorkbook workbook = null;
        private List<string> _sheetNames = new List<string>();
        private Dictionary<string, ExcelSheet> _sheetMap = new Dictionary<string, ExcelSheet>();

        public List<string> sheetNames
        {
            get
            {
                return _sheetNames;
            }
        }

        public string filePath
        {
            get
            {
                return _filePath;
            }
        }

        public void Load(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                _filePath = filePath;
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
                if (_sheetMap.ContainsKey(sheet.SheetName))
                {
                    return _sheetMap[sheet.SheetName];
                }
                ExcelSheet sheetWrapper = new ExcelSheet();
                sheetWrapper.Load(sheet);
                _sheetMap[sheet.SheetName] = sheetWrapper;
                return sheetWrapper;
            }
            return new ExcelSheet();
        }

        public void SaveSheet(VSheet sheet, string outputPath = null)
        {
            int sheetIndex = workbook.GetSheetIndex(sheet.name);
            if (sheetIndex == -1)
            {
                return;
            }
            ISheet worksheet = workbook.GetSheetAt(sheetIndex);
            if (worksheet != null)
            {
                sheet.Save(worksheet);
            }

            if ( string.IsNullOrEmpty(outputPath) )
            {
                outputPath = _filePath;
            }
            using (FileStream fs = File.OpenWrite(outputPath))
            {
                workbook.Write(fs);
            }
        }
    }
}
