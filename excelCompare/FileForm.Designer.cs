namespace excelCompare
{
    partial class FileForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.leftTextBox = new System.Windows.Forms.TextBox();
            this.rightTextBox = new System.Windows.Forms.TextBox();
            this.leftBrowseButton = new System.Windows.Forms.Button();
            this.rightBrowseButton = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "左侧";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 43);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "右侧";
            // 
            // leftTextBox
            // 
            this.leftTextBox.Location = new System.Drawing.Point(49, 13);
            this.leftTextBox.Name = "leftTextBox";
            this.leftTextBox.Size = new System.Drawing.Size(393, 21);
            this.leftTextBox.TabIndex = 2;
            this.leftTextBox.TextChanged += new System.EventHandler(this.OnTextChanged);
            // 
            // rightTextBox
            // 
            this.rightTextBox.Location = new System.Drawing.Point(48, 40);
            this.rightTextBox.Name = "rightTextBox";
            this.rightTextBox.Size = new System.Drawing.Size(393, 21);
            this.rightTextBox.TabIndex = 3;
            // 
            // leftBrowseButton
            // 
            this.leftBrowseButton.Location = new System.Drawing.Point(449, 12);
            this.leftBrowseButton.Name = "leftBrowseButton";
            this.leftBrowseButton.Size = new System.Drawing.Size(32, 23);
            this.leftBrowseButton.TabIndex = 4;
            this.leftBrowseButton.Text = "...";
            this.leftBrowseButton.UseVisualStyleBackColor = true;
            this.leftBrowseButton.Click += new System.EventHandler(this.OnBrowseLeftButtonClicked);
            // 
            // rightBrowseButton
            // 
            this.rightBrowseButton.Location = new System.Drawing.Point(449, 39);
            this.rightBrowseButton.Name = "rightBrowseButton";
            this.rightBrowseButton.Size = new System.Drawing.Size(32, 23);
            this.rightBrowseButton.TabIndex = 5;
            this.rightBrowseButton.Text = "...";
            this.rightBrowseButton.UseVisualStyleBackColor = true;
            this.rightBrowseButton.Click += new System.EventHandler(this.OnBrowseRightButtonClicked);
            // 
            // button3
            // 
            this.button3.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button3.Location = new System.Drawing.Point(365, 69);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 6;
            this.button3.Text = "取消";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // okButton
            // 
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(273, 69);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 23);
            this.okButton.TabIndex = 7;
            this.okButton.Text = "确定";
            this.okButton.UseVisualStyleBackColor = true;
            // 
            // FileForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(498, 104);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.rightBrowseButton);
            this.Controls.Add(this.leftBrowseButton);
            this.Controls.Add(this.rightTextBox);
            this.Controls.Add(this.leftTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FileForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "选择文件";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox leftTextBox;
        private System.Windows.Forms.TextBox rightTextBox;
        private System.Windows.Forms.Button leftBrowseButton;
        private System.Windows.Forms.Button rightBrowseButton;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button okButton;
    }
}