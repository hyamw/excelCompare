using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excelCompare
{
    internal class ExcelComparer
    {
        private List<SheetComparer> _sheets = new List<SheetComparer>();
        private Dictionary<string, SheetComparer> _nameMap = new Dictionary<string, SheetComparer>();
        private string[] _names = null;

        public string[] sheetNames
        {
            get
            {
                if (_names == null)
                {
                    _names = _nameMap.Keys.ToArray<string>();
                }
                return _names;
            }
        }

        public SheetComparer GetSheet(string name)
        {
            return _nameMap[name];
        }

        public void Load(string leftPath, string rightPath)
        {
            _names = null;
            _sheets.Clear();
            _nameMap.Clear();
            ExcelWorkbook leftWorkbook = new ExcelWorkbook();
            ExcelWorkbook rightWorkbook = new ExcelWorkbook();

            leftWorkbook.Load(leftPath);
            rightWorkbook.Load(rightPath);

            for (int i = 0; i < leftWorkbook.sheetNames.Count; i++)
            {
                string sheetName = leftWorkbook.sheetNames[i];
                SheetComparer sheet = new SheetComparer(sheetName);
                sheet.Execute(leftWorkbook, rightWorkbook);
                _nameMap.Add(sheetName, sheet);
            }

            for (int i = 0; i < rightWorkbook.sheetNames.Count; i++)
            {
                string sheetName = rightWorkbook.sheetNames[i];
                if (!_nameMap.ContainsKey(sheetName))
                {
                    SheetComparer sheet = new SheetComparer(sheetName);
                    sheet.Execute(leftWorkbook, rightWorkbook);
                    _nameMap.Add(sheetName, sheet);
                }
            }
        }
    }
}
