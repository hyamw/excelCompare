using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excelCompare
{
    public partial class FileForm : Form
    {
        public const string FILE_FILTER = "所有Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|Excel工作簿(*.xlsx)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls";
        private static readonly string[] FILTER_NAMES = new string[]
        {
            ".xlsx",
            ".xls"
        };
        public string leftPath
        {
            get
            {
                return leftTextBox.Text;
            }
        }

        public string rightPath
        {
            get
            {
                return rightTextBox.Text;
            }
        }
        public FileForm()
        {
            InitializeComponent();
            okButton.Enabled = false;
        }

        private void OnBrowseLeftButtonClicked(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = FILE_FILTER;
            if ( dialog.ShowDialog(this) == DialogResult.OK )
            {
                leftTextBox.Text = dialog.FileName;
                CheckInput();
            }
        }

        private void OnBrowseRightButtonClicked(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = FILE_FILTER;
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                rightTextBox.Text = dialog.FileName;
                CheckInput();
            }
        }

        private void OnTextChanged(object sender, EventArgs e)
        {
            CheckInput();
        }

        private void CheckInput()
        {
            bool allowStart = false;
            if (!string.IsNullOrEmpty(leftTextBox.Text) && !string.IsNullOrEmpty(rightTextBox.Text))
            {
                if (File.Exists(leftTextBox.Text) && File.Exists(rightTextBox.Text))
                {
                    allowStart = true;
                }
            }

            okButton.Enabled = allowStart;
        }

        public static int GetFilterIndex(string name)
        {
            string extension = Path.GetExtension(name).ToLower();
            for ( int i = 0; i < FILTER_NAMES.Length; i++ )
            {
                if ( FILTER_NAMES[i].CompareTo(extension) == 0 )
                {
                    return i + 2;
                }
            }
            return 0;
        }
    }
}
