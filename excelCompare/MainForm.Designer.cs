namespace excelCompare
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.operationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nextToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.previousToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.alignMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.compareToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.previousButton = new System.Windows.Forms.ToolStripButton();
            this.nextButton = new System.Windows.Forms.ToolStripButton();
            this.rootSplitContainer = new System.Windows.Forms.SplitContainer();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.leftSplitContainer = new System.Windows.Forms.SplitContainer();
            this.leftComboBox = new System.Windows.Forms.ComboBox();
            this.leftGrid = new System.Windows.Forms.DataGridView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.alignToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rightSplitContainer = new System.Windows.Forms.SplitContainer();
            this.rightComboBox = new System.Windows.Forms.ComboBox();
            this.rightGrid = new System.Windows.Forms.DataGridView();
            this.rowDiffGrid = new System.Windows.Forms.DataGridView();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rootSplitContainer)).BeginInit();
            this.rootSplitContainer.Panel1.SuspendLayout();
            this.rootSplitContainer.Panel2.SuspendLayout();
            this.rootSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.leftSplitContainer)).BeginInit();
            this.leftSplitContainer.Panel1.SuspendLayout();
            this.leftSplitContainer.Panel2.SuspendLayout();
            this.leftSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.leftGrid)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rightSplitContainer)).BeginInit();
            this.rightSplitContainer.Panel1.SuspendLayout();
            this.rightSplitContainer.Panel2.SuspendLayout();
            this.rightSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rightGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rowDiffGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.operationToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1309, 25);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // operationToolStripMenuItem
            // 
            this.operationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.nextToolStripMenuItem,
            this.previousToolStripMenuItem,
            this.alignMenuItem});
            this.operationToolStripMenuItem.Name = "operationToolStripMenuItem";
            this.operationToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.operationToolStripMenuItem.Text = "操作(&T)";
            this.operationToolStripMenuItem.DropDownOpening += new System.EventHandler(this.OnOperationDropDownOpening);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.openToolStripMenuItem.Text = "打开(&O)";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.OnOpenClicked);
            // 
            // nextToolStripMenuItem
            // 
            this.nextToolStripMenuItem.Name = "nextToolStripMenuItem";
            this.nextToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.nextToolStripMenuItem.Text = "下一个差异(&N)";
            // 
            // previousToolStripMenuItem
            // 
            this.previousToolStripMenuItem.Name = "previousToolStripMenuItem";
            this.previousToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.previousToolStripMenuItem.Text = "上一个差异(&P)";
            // 
            // alignMenuItem
            // 
            this.alignMenuItem.Name = "alignMenuItem";
            this.alignMenuItem.Size = new System.Drawing.Size(154, 22);
            this.alignMenuItem.Text = "对齐(&A)";
            this.alignMenuItem.Click += new System.EventHandler(this.OnAlignMenuClicked);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.compareToolStripButton,
            this.previousButton,
            this.nextButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 25);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1309, 25);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // compareToolStripButton
            // 
            this.compareToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.compareToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("compareToolStripButton.Image")));
            this.compareToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.compareToolStripButton.Name = "compareToolStripButton";
            this.compareToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.compareToolStripButton.Text = "选择表格对比";
            this.compareToolStripButton.ToolTipText = "选择表格对比";
            this.compareToolStripButton.Click += new System.EventHandler(this.OnCompareButtonClicked);
            // 
            // previousButton
            // 
            this.previousButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.previousButton.Image = ((System.Drawing.Image)(resources.GetObject("previousButton.Image")));
            this.previousButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.previousButton.Name = "previousButton";
            this.previousButton.Size = new System.Drawing.Size(23, 22);
            this.previousButton.Text = "toolStripButton1";
            this.previousButton.ToolTipText = "上一个差异";
            this.previousButton.Click += new System.EventHandler(this.OnPreviousButtonClicked);
            // 
            // nextButton
            // 
            this.nextButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.nextButton.Image = ((System.Drawing.Image)(resources.GetObject("nextButton.Image")));
            this.nextButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(23, 22);
            this.nextButton.Text = "Next";
            this.nextButton.ToolTipText = "下一个差异";
            this.nextButton.Click += new System.EventHandler(this.OnNextButtonClicked);
            // 
            // rootSplitContainer
            // 
            this.rootSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rootSplitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.rootSplitContainer.IsSplitterFixed = true;
            this.rootSplitContainer.Location = new System.Drawing.Point(0, 50);
            this.rootSplitContainer.Name = "rootSplitContainer";
            this.rootSplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // rootSplitContainer.Panel1
            // 
            this.rootSplitContainer.Panel1.Controls.Add(this.splitContainer1);
            // 
            // rootSplitContainer.Panel2
            // 
            this.rootSplitContainer.Panel2.Controls.Add(this.rowDiffGrid);
            this.rootSplitContainer.Size = new System.Drawing.Size(1309, 590);
            this.rootSplitContainer.SplitterDistance = 500;
            this.rootSplitContainer.TabIndex = 2;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.leftSplitContainer);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.rightSplitContainer);
            this.splitContainer1.Size = new System.Drawing.Size(1309, 500);
            this.splitContainer1.SplitterDistance = 655;
            this.splitContainer1.TabIndex = 0;
            // 
            // leftSplitContainer
            // 
            this.leftSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.leftSplitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.leftSplitContainer.IsSplitterFixed = true;
            this.leftSplitContainer.Location = new System.Drawing.Point(0, 0);
            this.leftSplitContainer.Name = "leftSplitContainer";
            this.leftSplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // leftSplitContainer.Panel1
            // 
            this.leftSplitContainer.Panel1.Controls.Add(this.leftComboBox);
            this.leftSplitContainer.Panel1MinSize = 20;
            // 
            // leftSplitContainer.Panel2
            // 
            this.leftSplitContainer.Panel2.Controls.Add(this.leftGrid);
            this.leftSplitContainer.Size = new System.Drawing.Size(655, 500);
            this.leftSplitContainer.SplitterDistance = 25;
            this.leftSplitContainer.TabIndex = 0;
            // 
            // leftComboBox
            // 
            this.leftComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.leftComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.leftComboBox.FormattingEnabled = true;
            this.leftComboBox.Location = new System.Drawing.Point(0, 0);
            this.leftComboBox.Name = "leftComboBox";
            this.leftComboBox.Size = new System.Drawing.Size(655, 20);
            this.leftComboBox.TabIndex = 0;
            this.leftComboBox.SelectedValueChanged += new System.EventHandler(this.OnSheetSelectionChanged);
            // 
            // leftGrid
            // 
            this.leftGrid.AllowUserToAddRows = false;
            this.leftGrid.AllowUserToDeleteRows = false;
            this.leftGrid.AllowUserToResizeRows = false;
            this.leftGrid.BackgroundColor = System.Drawing.Color.White;
            this.leftGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.leftGrid.ContextMenuStrip = this.contextMenuStrip1;
            this.leftGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.leftGrid.Location = new System.Drawing.Point(0, 0);
            this.leftGrid.MultiSelect = false;
            this.leftGrid.Name = "leftGrid";
            this.leftGrid.ReadOnly = true;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black;
            this.leftGrid.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.leftGrid.RowTemplate.Height = 23;
            this.leftGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.leftGrid.Size = new System.Drawing.Size(655, 471);
            this.leftGrid.TabIndex = 1;
            this.leftGrid.RowHeadersWidthChanged += new System.EventHandler(this.OnRowHeaderWidthChanged);
            this.leftGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.OnCellClicked);
            this.leftGrid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.OnCellFormatting);
            this.leftGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.OnCellPainting);
            this.leftGrid.ColumnWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.OnColumnWidthChanged);
            this.leftGrid.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(this.OnRowStateChanged);
            this.leftGrid.Scroll += new System.Windows.Forms.ScrollEventHandler(this.OnGridViewScroll);
            this.leftGrid.SelectionChanged += new System.EventHandler(this.OnGridViewSelectionChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.alignToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(117, 26);
            // 
            // alignToolStripMenuItem
            // 
            this.alignToolStripMenuItem.Name = "alignToolStripMenuItem";
            this.alignToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
            this.alignToolStripMenuItem.Text = "对齐(&A)";
            this.alignToolStripMenuItem.Click += new System.EventHandler(this.OnAlignMenuClicked);
            // 
            // rightSplitContainer
            // 
            this.rightSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rightSplitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.rightSplitContainer.IsSplitterFixed = true;
            this.rightSplitContainer.Location = new System.Drawing.Point(0, 0);
            this.rightSplitContainer.Name = "rightSplitContainer";
            this.rightSplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // rightSplitContainer.Panel1
            // 
            this.rightSplitContainer.Panel1.Controls.Add(this.rightComboBox);
            this.rightSplitContainer.Panel1MinSize = 20;
            // 
            // rightSplitContainer.Panel2
            // 
            this.rightSplitContainer.Panel2.Controls.Add(this.rightGrid);
            this.rightSplitContainer.Size = new System.Drawing.Size(650, 500);
            this.rightSplitContainer.SplitterDistance = 25;
            this.rightSplitContainer.TabIndex = 0;
            // 
            // rightComboBox
            // 
            this.rightComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rightComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.rightComboBox.FormattingEnabled = true;
            this.rightComboBox.Location = new System.Drawing.Point(0, 0);
            this.rightComboBox.Name = "rightComboBox";
            this.rightComboBox.Size = new System.Drawing.Size(650, 20);
            this.rightComboBox.TabIndex = 0;
            this.rightComboBox.SelectedValueChanged += new System.EventHandler(this.OnSheetSelectionChanged);
            // 
            // rightGrid
            // 
            this.rightGrid.AllowUserToAddRows = false;
            this.rightGrid.AllowUserToDeleteRows = false;
            this.rightGrid.AllowUserToResizeRows = false;
            this.rightGrid.BackgroundColor = System.Drawing.Color.White;
            this.rightGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.rightGrid.ContextMenuStrip = this.contextMenuStrip1;
            this.rightGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rightGrid.Location = new System.Drawing.Point(0, 0);
            this.rightGrid.MultiSelect = false;
            this.rightGrid.Name = "rightGrid";
            this.rightGrid.ReadOnly = true;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            this.rightGrid.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.rightGrid.RowTemplate.Height = 23;
            this.rightGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.rightGrid.Size = new System.Drawing.Size(650, 471);
            this.rightGrid.TabIndex = 1;
            this.rightGrid.RowHeadersWidthChanged += new System.EventHandler(this.OnRowHeaderWidthChanged);
            this.rightGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.OnCellClicked);
            this.rightGrid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.OnCellFormatting);
            this.rightGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.OnCellPainting);
            this.rightGrid.ColumnWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.OnColumnWidthChanged);
            this.rightGrid.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(this.OnRowStateChanged);
            this.rightGrid.Scroll += new System.Windows.Forms.ScrollEventHandler(this.OnGridViewScroll);
            this.rightGrid.SelectionChanged += new System.EventHandler(this.OnGridViewSelectionChanged);
            // 
            // rowDiffGrid
            // 
            this.rowDiffGrid.AllowUserToAddRows = false;
            this.rowDiffGrid.AllowUserToDeleteRows = false;
            this.rowDiffGrid.AllowUserToResizeRows = false;
            this.rowDiffGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.rowDiffGrid.BackgroundColor = System.Drawing.Color.White;
            this.rowDiffGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.rowDiffGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rowDiffGrid.Location = new System.Drawing.Point(0, 0);
            this.rowDiffGrid.MultiSelect = false;
            this.rowDiffGrid.Name = "rowDiffGrid";
            this.rowDiffGrid.RowHeadersVisible = false;
            this.rowDiffGrid.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption;
            this.rowDiffGrid.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.rowDiffGrid.RowTemplate.Height = 23;
            this.rowDiffGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.rowDiffGrid.Size = new System.Drawing.Size(1309, 86);
            this.rowDiffGrid.TabIndex = 0;
            this.rowDiffGrid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.OnRowGridCellFormatting);
            this.rowDiffGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.OnRowGridCellPainting);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1309, 640);
            this.Controls.Add(this.rootSplitContainer);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.Name = "MainForm";
            this.Text = "Excel对比工具";
            this.KeyUp += new System.Windows.Forms.KeyEventHandler(this.OnKeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.rootSplitContainer.Panel1.ResumeLayout(false);
            this.rootSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.rootSplitContainer)).EndInit();
            this.rootSplitContainer.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.leftSplitContainer.Panel1.ResumeLayout(false);
            this.leftSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.leftSplitContainer)).EndInit();
            this.leftSplitContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.leftGrid)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.rightSplitContainer.Panel1.ResumeLayout(false);
            this.rightSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.rightSplitContainer)).EndInit();
            this.rightSplitContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.rightGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rowDiffGrid)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem operationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem nextToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem previousToolStripMenuItem;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton previousButton;
        private System.Windows.Forms.ToolStripButton nextButton;
        private System.Windows.Forms.SplitContainer rootSplitContainer;
        private System.Windows.Forms.DataGridView rowDiffGrid;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem alignToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem alignMenuItem;
        private System.Windows.Forms.SplitContainer leftSplitContainer;
        private System.Windows.Forms.DataGridView leftGrid;
        private System.Windows.Forms.ComboBox leftComboBox;
        private System.Windows.Forms.SplitContainer rightSplitContainer;
        private System.Windows.Forms.DataGridView rightGrid;
        private System.Windows.Forms.ComboBox rightComboBox;
        private System.Windows.Forms.ToolStripButton compareToolStripButton;
    }
}

