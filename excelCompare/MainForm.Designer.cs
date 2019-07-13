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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.operationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.nextToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.previousToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copy2LeftMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copy2RightMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.alignMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.compareToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.previousButton = new System.Windows.Forms.ToolStripButton();
            this.nextButton = new System.Windows.Forms.ToolStripButton();
            this.copyToolstripButton = new System.Windows.Forms.ToolStripButton();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.alignToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.leftToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.rightToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mainSplitContainer = new System.Windows.Forms.SplitContainer();
            this.rootSplitContainer = new System.Windows.Forms.SplitContainer();
            this.topSplitContainer = new System.Windows.Forms.SplitContainer();
            this.leftSplitContainer = new System.Windows.Forms.SplitContainer();
            this.panel2 = new System.Windows.Forms.Panel();
            this.leftComboBox = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.leftFileTextBox = new System.Windows.Forms.TextBox();
            this.rightSplitContainer = new System.Windows.Forms.SplitContainer();
            this.panel4 = new System.Windows.Forms.Panel();
            this.rightComboBox = new System.Windows.Forms.ComboBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.rightFileTextBox = new System.Windows.Forms.TextBox();
            this.rowDiffGrid = new System.Windows.Forms.DataGridView();
            this.differenceScrollBar = new excelCompare.DifferenceScrollBar();
            this.leftGrid = new excelCompare.CustomDataGridView();
            this.rightGrid = new excelCompare.CustomDataGridView();
            this.menuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.mainSplitContainer)).BeginInit();
            this.mainSplitContainer.Panel1.SuspendLayout();
            this.mainSplitContainer.Panel2.SuspendLayout();
            this.mainSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rootSplitContainer)).BeginInit();
            this.rootSplitContainer.Panel1.SuspendLayout();
            this.rootSplitContainer.Panel2.SuspendLayout();
            this.rootSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.topSplitContainer)).BeginInit();
            this.topSplitContainer.Panel1.SuspendLayout();
            this.topSplitContainer.Panel2.SuspendLayout();
            this.topSplitContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.leftSplitContainer)).BeginInit();
            this.leftSplitContainer.Panel1.SuspendLayout();
            this.leftSplitContainer.Panel2.SuspendLayout();
            this.leftSplitContainer.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rightSplitContainer)).BeginInit();
            this.rightSplitContainer.Panel1.SuspendLayout();
            this.rightSplitContainer.Panel2.SuspendLayout();
            this.rightSplitContainer.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rowDiffGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.leftGrid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rightGrid)).BeginInit();
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
            this.copy2LeftMenuItem,
            this.copy2RightMenuItem,
            this.alignMenuItem,
            this.toolStripSeparator1,
            this.saveToolStripMenuItem,
            this.saveAsToolStripMenuItem,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
            this.operationToolStripMenuItem.Name = "operationToolStripMenuItem";
            this.operationToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.operationToolStripMenuItem.Text = "操作(&T)";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.openToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.openToolStripMenuItem.Text = "打开(&O)";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.OnOpenClicked);
            // 
            // nextToolStripMenuItem
            // 
            this.nextToolStripMenuItem.Enabled = false;
            this.nextToolStripMenuItem.Name = "nextToolStripMenuItem";
            this.nextToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.nextToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.nextToolStripMenuItem.Text = "下一个差异";
            this.nextToolStripMenuItem.Click += new System.EventHandler(this.OnNextButtonClicked);
            // 
            // previousToolStripMenuItem
            // 
            this.previousToolStripMenuItem.Enabled = false;
            this.previousToolStripMenuItem.Name = "previousToolStripMenuItem";
            this.previousToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
            this.previousToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.previousToolStripMenuItem.Text = "上一个差异";
            this.previousToolStripMenuItem.Click += new System.EventHandler(this.OnPreviousButtonClicked);
            // 
            // copy2LeftMenuItem
            // 
            this.copy2LeftMenuItem.Enabled = false;
            this.copy2LeftMenuItem.Name = "copy2LeftMenuItem";
            this.copy2LeftMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.L)));
            this.copy2LeftMenuItem.Size = new System.Drawing.Size(190, 22);
            this.copy2LeftMenuItem.Text = "复制到左边";
            this.copy2LeftMenuItem.Click += new System.EventHandler(this.OnCopy2LeftClicked);
            // 
            // copy2RightMenuItem
            // 
            this.copy2RightMenuItem.Enabled = false;
            this.copy2RightMenuItem.Name = "copy2RightMenuItem";
            this.copy2RightMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.R)));
            this.copy2RightMenuItem.Size = new System.Drawing.Size(190, 22);
            this.copy2RightMenuItem.Text = "复制到右边";
            this.copy2RightMenuItem.Click += new System.EventHandler(this.OnCopy2RightClicked);
            // 
            // alignMenuItem
            // 
            this.alignMenuItem.Enabled = false;
            this.alignMenuItem.Name = "alignMenuItem";
            this.alignMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F7;
            this.alignMenuItem.Size = new System.Drawing.Size(190, 22);
            this.alignMenuItem.Text = "对齐";
            this.alignMenuItem.Click += new System.EventHandler(this.OnAlignMenuClicked);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(187, 6);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Enabled = false;
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.saveToolStripMenuItem.Text = "保存";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.OnSaveMenuClicked);
            // 
            // saveAsToolStripMenuItem
            // 
            this.saveAsToolStripMenuItem.Enabled = false;
            this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
            this.saveAsToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift) 
            | System.Windows.Forms.Keys.S)));
            this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.saveAsToolStripMenuItem.Text = "另存为";
            this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.OnSaveAsMenuClicked);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(187, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(190, 22);
            this.exitToolStripMenuItem.Text = "退出(&E)";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.OnExitMenuClicked);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.compareToolStripButton,
            this.previousButton,
            this.nextButton,
            this.copyToolstripButton});
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
            this.compareToolStripButton.ToolTipText = "选择表格对比(F7)";
            this.compareToolStripButton.Click += new System.EventHandler(this.OnCompareButtonClicked);
            // 
            // previousButton
            // 
            this.previousButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.previousButton.Enabled = false;
            this.previousButton.Image = ((System.Drawing.Image)(resources.GetObject("previousButton.Image")));
            this.previousButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.previousButton.Name = "previousButton";
            this.previousButton.Size = new System.Drawing.Size(23, 22);
            this.previousButton.Text = "toolStripButton1";
            this.previousButton.ToolTipText = "上一个差异(Ctrl+P)";
            this.previousButton.Click += new System.EventHandler(this.OnPreviousButtonClicked);
            // 
            // nextButton
            // 
            this.nextButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.nextButton.Enabled = false;
            this.nextButton.Image = ((System.Drawing.Image)(resources.GetObject("nextButton.Image")));
            this.nextButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.nextButton.Name = "nextButton";
            this.nextButton.Size = new System.Drawing.Size(23, 22);
            this.nextButton.Text = "Next";
            this.nextButton.ToolTipText = "下一个差异(Ctrl+N)";
            this.nextButton.Click += new System.EventHandler(this.OnNextButtonClicked);
            // 
            // copyToolstripButton
            // 
            this.copyToolstripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.copyToolstripButton.Enabled = false;
            this.copyToolstripButton.Image = global::excelCompare.Properties.Resources.left;
            this.copyToolstripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.copyToolstripButton.Name = "copyToolstripButton";
            this.copyToolstripButton.Size = new System.Drawing.Size(23, 22);
            this.copyToolstripButton.Text = "复制到左边";
            this.copyToolstripButton.ToolTipText = "复制到左边(Ctrl+L)";
            this.copyToolstripButton.Click += new System.EventHandler(this.OnCopyButtonClicked);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.alignToolStripMenuItem,
            this.leftToolStripMenuItem,
            this.rightToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(153, 70);
            // 
            // alignToolStripMenuItem
            // 
            this.alignToolStripMenuItem.Name = "alignToolStripMenuItem";
            this.alignToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.alignToolStripMenuItem.Text = "对齐(&A)";
            this.alignToolStripMenuItem.Click += new System.EventHandler(this.OnAlignMenuClicked);
            // 
            // leftToolStripMenuItem
            // 
            this.leftToolStripMenuItem.Name = "leftToolStripMenuItem";
            this.leftToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.leftToolStripMenuItem.Text = "复制到左边(&L)";
            this.leftToolStripMenuItem.Click += new System.EventHandler(this.OnCopy2LeftClicked);
            // 
            // rightToolStripMenuItem
            // 
            this.rightToolStripMenuItem.Name = "rightToolStripMenuItem";
            this.rightToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.rightToolStripMenuItem.Text = "复制到右边(&R)";
            this.rightToolStripMenuItem.Click += new System.EventHandler(this.OnCopy2RightClicked);
            // 
            // mainSplitContainer
            // 
            this.mainSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainSplitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.mainSplitContainer.IsSplitterFixed = true;
            this.mainSplitContainer.Location = new System.Drawing.Point(0, 50);
            this.mainSplitContainer.Name = "mainSplitContainer";
            // 
            // mainSplitContainer.Panel1
            // 
            this.mainSplitContainer.Panel1.Controls.Add(this.differenceScrollBar);
            // 
            // mainSplitContainer.Panel2
            // 
            this.mainSplitContainer.Panel2.Controls.Add(this.rootSplitContainer);
            this.mainSplitContainer.Size = new System.Drawing.Size(1309, 590);
            this.mainSplitContainer.SplitterDistance = 52;
            this.mainSplitContainer.TabIndex = 2;
            // 
            // rootSplitContainer
            // 
            this.rootSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rootSplitContainer.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.rootSplitContainer.IsSplitterFixed = true;
            this.rootSplitContainer.Location = new System.Drawing.Point(0, 0);
            this.rootSplitContainer.Name = "rootSplitContainer";
            this.rootSplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // rootSplitContainer.Panel1
            // 
            this.rootSplitContainer.Panel1.Controls.Add(this.topSplitContainer);
            // 
            // rootSplitContainer.Panel2
            // 
            this.rootSplitContainer.Panel2.Controls.Add(this.rowDiffGrid);
            this.rootSplitContainer.Size = new System.Drawing.Size(1253, 590);
            this.rootSplitContainer.SplitterDistance = 500;
            this.rootSplitContainer.TabIndex = 3;
            // 
            // topSplitContainer
            // 
            this.topSplitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.topSplitContainer.Location = new System.Drawing.Point(0, 0);
            this.topSplitContainer.Name = "topSplitContainer";
            // 
            // topSplitContainer.Panel1
            // 
            this.topSplitContainer.Panel1.Controls.Add(this.leftSplitContainer);
            // 
            // topSplitContainer.Panel2
            // 
            this.topSplitContainer.Panel2.Controls.Add(this.rightSplitContainer);
            this.topSplitContainer.Size = new System.Drawing.Size(1253, 500);
            this.topSplitContainer.SplitterDistance = 625;
            this.topSplitContainer.TabIndex = 0;
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
            this.leftSplitContainer.Panel1.Controls.Add(this.panel2);
            this.leftSplitContainer.Panel1.Controls.Add(this.panel1);
            this.leftSplitContainer.Panel1MinSize = 20;
            // 
            // leftSplitContainer.Panel2
            // 
            this.leftSplitContainer.Panel2.Controls.Add(this.leftGrid);
            this.leftSplitContainer.Size = new System.Drawing.Size(625, 500);
            this.leftSplitContainer.SplitterDistance = 43;
            this.leftSplitContainer.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.leftComboBox);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 22);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(625, 21);
            this.panel2.TabIndex = 1;
            // 
            // leftComboBox
            // 
            this.leftComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.leftComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.leftComboBox.FormattingEnabled = true;
            this.leftComboBox.Location = new System.Drawing.Point(0, 0);
            this.leftComboBox.Name = "leftComboBox";
            this.leftComboBox.Size = new System.Drawing.Size(625, 20);
            this.leftComboBox.TabIndex = 2;
            this.leftComboBox.SelectedValueChanged += new System.EventHandler(this.OnSheetSelectionChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.leftFileTextBox);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(625, 22);
            this.panel1.TabIndex = 0;
            // 
            // leftFileTextBox
            // 
            this.leftFileTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.leftFileTextBox.Location = new System.Drawing.Point(0, 0);
            this.leftFileTextBox.Name = "leftFileTextBox";
            this.leftFileTextBox.ReadOnly = true;
            this.leftFileTextBox.Size = new System.Drawing.Size(625, 21);
            this.leftFileTextBox.TabIndex = 18;
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
            this.rightSplitContainer.Panel1.Controls.Add(this.panel4);
            this.rightSplitContainer.Panel1.Controls.Add(this.panel3);
            this.rightSplitContainer.Panel1MinSize = 20;
            // 
            // rightSplitContainer.Panel2
            // 
            this.rightSplitContainer.Panel2.Controls.Add(this.rightGrid);
            this.rightSplitContainer.Size = new System.Drawing.Size(624, 500);
            this.rightSplitContainer.SplitterDistance = 43;
            this.rightSplitContainer.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.rightComboBox);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel4.Location = new System.Drawing.Point(0, 22);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(624, 21);
            this.panel4.TabIndex = 1;
            // 
            // rightComboBox
            // 
            this.rightComboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rightComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.rightComboBox.FormattingEnabled = true;
            this.rightComboBox.Location = new System.Drawing.Point(0, 0);
            this.rightComboBox.Name = "rightComboBox";
            this.rightComboBox.Size = new System.Drawing.Size(624, 20);
            this.rightComboBox.TabIndex = 1;
            this.rightComboBox.SelectedValueChanged += new System.EventHandler(this.OnSheetSelectionChanged);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.rightFileTextBox);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(624, 22);
            this.panel3.TabIndex = 0;
            // 
            // rightFileTextBox
            // 
            this.rightFileTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rightFileTextBox.Location = new System.Drawing.Point(0, 0);
            this.rightFileTextBox.Name = "rightFileTextBox";
            this.rightFileTextBox.ReadOnly = true;
            this.rightFileTextBox.Size = new System.Drawing.Size(624, 21);
            this.rightFileTextBox.TabIndex = 1;
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
            this.rowDiffGrid.Size = new System.Drawing.Size(1253, 86);
            this.rowDiffGrid.TabIndex = 0;
            this.rowDiffGrid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.OnRowGridCellFormatting);
            this.rowDiffGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.OnRowGridCellPainting);
            // 
            // differenceScrollBar
            // 
            this.differenceScrollBar.ArrowColor = System.Drawing.Color.FromArgb(((int)(((byte)(65)))), ((int)(((byte)(65)))), ((int)(((byte)(255)))));
            this.differenceScrollBar.BackColor = System.Drawing.SystemColors.Control;
            this.differenceScrollBar.BorderColor = System.Drawing.SystemColors.Control;
            this.differenceScrollBar.DataSource = null;
            this.differenceScrollBar.DiffColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(65)))), ((int)(((byte)(65)))));
            this.differenceScrollBar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.differenceScrollBar.Location = new System.Drawing.Point(0, 0);
            this.differenceScrollBar.Name = "differenceScrollBar";
            this.differenceScrollBar.Selection = -1;
            this.differenceScrollBar.Size = new System.Drawing.Size(52, 590);
            this.differenceScrollBar.TabIndex = 0;
            this.differenceScrollBar.Text = "differenceScrollBar1";
            this.differenceScrollBar.Scroll += new System.Windows.Forms.ScrollEventHandler(this.OnDifferenceScrolled);
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
            this.leftGrid.Name = "leftGrid";
            this.leftGrid.ReadOnly = true;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black;
            this.leftGrid.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.leftGrid.RowTemplate.Height = 23;
            this.leftGrid.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.leftGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.leftGrid.Size = new System.Drawing.Size(625, 453);
            this.leftGrid.TabIndex = 1;
            this.leftGrid.RowHeadersWidthChanged += new System.EventHandler(this.OnRowHeaderWidthChanged);
            this.leftGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.OnCellClicked);
            this.leftGrid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.OnCellFormatting);
            this.leftGrid.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.OnCellMouseDown);
            this.leftGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.OnCellPainting);
            this.leftGrid.ColumnWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.OnColumnWidthChanged);
            this.leftGrid.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(this.OnRowStateChanged);
            this.leftGrid.Scroll += new System.Windows.Forms.ScrollEventHandler(this.OnGridViewScroll);
            this.leftGrid.SelectionChanged += new System.EventHandler(this.OnGridViewSelectionChanged);
            this.leftGrid.Enter += new System.EventHandler(this.OnGridViewFocusEnter);
            this.leftGrid.Leave += new System.EventHandler(this.OnGridViewFocusLeave);
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
            this.rightGrid.Name = "rightGrid";
            this.rightGrid.ReadOnly = true;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            this.rightGrid.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.rightGrid.RowTemplate.Height = 23;
            this.rightGrid.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.rightGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.rightGrid.Size = new System.Drawing.Size(624, 453);
            this.rightGrid.TabIndex = 1;
            this.rightGrid.RowHeadersWidthChanged += new System.EventHandler(this.OnRowHeaderWidthChanged);
            this.rightGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.OnCellClicked);
            this.rightGrid.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.OnCellFormatting);
            this.rightGrid.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.OnCellMouseDown);
            this.rightGrid.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.OnCellPainting);
            this.rightGrid.ColumnWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.OnColumnWidthChanged);
            this.rightGrid.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(this.OnRowStateChanged);
            this.rightGrid.Scroll += new System.Windows.Forms.ScrollEventHandler(this.OnGridViewScroll);
            this.rightGrid.SelectionChanged += new System.EventHandler(this.OnGridViewSelectionChanged);
            this.rightGrid.Enter += new System.EventHandler(this.OnGridViewFocusEnter);
            this.rightGrid.Leave += new System.EventHandler(this.OnGridViewFocusLeave);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1309, 640);
            this.Controls.Add(this.mainSplitContainer);
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
            this.contextMenuStrip1.ResumeLayout(false);
            this.mainSplitContainer.Panel1.ResumeLayout(false);
            this.mainSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.mainSplitContainer)).EndInit();
            this.mainSplitContainer.ResumeLayout(false);
            this.rootSplitContainer.Panel1.ResumeLayout(false);
            this.rootSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.rootSplitContainer)).EndInit();
            this.rootSplitContainer.ResumeLayout(false);
            this.topSplitContainer.Panel1.ResumeLayout(false);
            this.topSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.topSplitContainer)).EndInit();
            this.topSplitContainer.ResumeLayout(false);
            this.leftSplitContainer.Panel1.ResumeLayout(false);
            this.leftSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.leftSplitContainer)).EndInit();
            this.leftSplitContainer.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.rightSplitContainer.Panel1.ResumeLayout(false);
            this.rightSplitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.rightSplitContainer)).EndInit();
            this.rightSplitContainer.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rowDiffGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.leftGrid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rightGrid)).EndInit();
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
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem alignToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem alignMenuItem;
        private System.Windows.Forms.ToolStripButton compareToolStripButton;
        private System.Windows.Forms.ToolStripButton copyToolstripButton;
        private System.Windows.Forms.ToolStripMenuItem leftToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rightToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem saveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copy2LeftMenuItem;
        private System.Windows.Forms.ToolStripMenuItem copy2RightMenuItem;
        private System.Windows.Forms.SplitContainer mainSplitContainer;
        private System.Windows.Forms.SplitContainer rootSplitContainer;
        private System.Windows.Forms.SplitContainer topSplitContainer;
        private System.Windows.Forms.SplitContainer leftSplitContainer;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ComboBox leftComboBox;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox leftFileTextBox;
        private CustomDataGridView leftGrid;
        private System.Windows.Forms.SplitContainer rightSplitContainer;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.ComboBox rightComboBox;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox rightFileTextBox;
        private CustomDataGridView rightGrid;
        private System.Windows.Forms.DataGridView rowDiffGrid;
        private DifferenceScrollBar differenceScrollBar;
    }
}

