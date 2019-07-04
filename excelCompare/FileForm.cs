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
            dialog.Filter = "所有Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|Excel工作簿(*.xls)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls";
            if ( dialog.ShowDialog(this) == DialogResult.OK )
            {
                leftTextBox.Text = dialog.FileName;
                CheckInput();
            }
        }

        private void OnBrowseRightButtonClicked(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "所有Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx|Excel工作簿(*.xls)|*.xlsx|Excel 97-2003 工作簿(*.xls)|*.xls";
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
    }
}
