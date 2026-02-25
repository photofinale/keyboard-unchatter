namespace KeyboardUnchatter
{
    partial class MainWindow
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this._statusPanel = new System.Windows.Forms.Panel();
            this.labelTypingSpeed = new System.Windows.Forms.Label();
            this._statusPanelLabel = new System.Windows.Forms.Label();
            this._buttonActivate = new System.Windows.Forms.Button();
            this._mainDataGrid = new System.Windows.Forms.DataGridView();
            this.Key = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PressCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ChatterCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FailureRate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this._dataGridGroupBox = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbTestInput = new System.Windows.Forms.RichTextBox();
            this.listBoxIntervals = new System.Windows.Forms.ListBox();
            this._runAtStartupCheckBox = new System.Windows.Forms.CheckBox();
            this._startAsAdminCheckBox = new System.Windows.Forms.CheckBox();
            this._activateOnLaunchCheckBox = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this._minimizeCheckBox = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this._thresholdTimeInput = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this._notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this._notifyIconMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this._openMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this._exitMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnRelaunchAsAdmin = new System.Windows.Forms.Button();
            this._statusPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this._mainDataGrid)).BeginInit();
            this._dataGridGroupBox.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this._thresholdTimeInput)).BeginInit();
            this._notifyIconMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // _statusPanel
            // 
            this._statusPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._statusPanel.BackColor = System.Drawing.SystemColors.ControlDark;
            this._statusPanel.Controls.Add(this.labelTypingSpeed);
            this._statusPanel.Controls.Add(this._statusPanelLabel);
            this._statusPanel.Location = new System.Drawing.Point(0, 1);
            this._statusPanel.Name = "_statusPanel";
            this._statusPanel.Size = new System.Drawing.Size(640, 41);
            this._statusPanel.TabIndex = 0;
            // 
            // labelTypingSpeed
            // 
            this.labelTypingSpeed.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelTypingSpeed.AutoSize = true;
            this.labelTypingSpeed.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F);
            this.labelTypingSpeed.Location = new System.Drawing.Point(161, 4);
            this.labelTypingSpeed.Name = "labelTypingSpeed";
            this.labelTypingSpeed.Size = new System.Drawing.Size(286, 37);
            this.labelTypingSpeed.TabIndex = 12;
            this.labelTypingSpeed.Text = "Typing Speed: N/A";
            // 
            // _statusPanelLabel
            // 
            this._statusPanelLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._statusPanelLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this._statusPanelLabel.Location = new System.Drawing.Point(0, 4);
            this._statusPanelLabel.Name = "_statusPanelLabel";
            this._statusPanelLabel.Size = new System.Drawing.Size(640, 37);
            this._statusPanelLabel.TabIndex = 0;
            this._statusPanelLabel.Text = "Not Active";
            this._statusPanelLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // _buttonActivate
            // 
            this._buttonActivate.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this._buttonActivate.Location = new System.Drawing.Point(248, 471);
            this._buttonActivate.Name = "_buttonActivate";
            this._buttonActivate.Size = new System.Drawing.Size(137, 27);
            this._buttonActivate.TabIndex = 2;
            this._buttonActivate.TabStop = false;
            this._buttonActivate.Text = "Activate";
            this._buttonActivate.UseVisualStyleBackColor = true;
            this._buttonActivate.Click += new System.EventHandler(this.OnActivateButtonClick);
            // 
            // _mainDataGrid
            // 
            this._mainDataGrid.AllowUserToAddRows = false;
            this._mainDataGrid.AllowUserToDeleteRows = false;
            this._mainDataGrid.AllowUserToResizeRows = false;
            this._mainDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._mainDataGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this._mainDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this._mainDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Key,
            this.PressCount,
            this.ChatterCount,
            this.FailureRate});
            this._mainDataGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this._mainDataGrid.Location = new System.Drawing.Point(6, 19);
            this._mainDataGrid.MultiSelect = false;
            this._mainDataGrid.Name = "_mainDataGrid";
            this._mainDataGrid.RowHeadersVisible = false;
            this._mainDataGrid.Size = new System.Drawing.Size(603, 239);
            this._mainDataGrid.TabIndex = 3;
            this._mainDataGrid.SortCompare += new System.Windows.Forms.DataGridViewSortCompareEventHandler(this.DataSortCompare);
            // 
            // Key
            // 
            this.Key.HeaderText = "Key";
            this.Key.Name = "Key";
            // 
            // PressCount
            // 
            this.PressCount.HeaderText = "Press Count";
            this.PressCount.Name = "PressCount";
            // 
            // ChatterCount
            // 
            this.ChatterCount.HeaderText = "Chatter Count";
            this.ChatterCount.Name = "ChatterCount";
            // 
            // FailureRate
            // 
            this.FailureRate.HeaderText = "Failure Rate";
            this.FailureRate.Name = "FailureRate";
            // 
            // _dataGridGroupBox
            // 
            this._dataGridGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this._dataGridGroupBox.Controls.Add(this._mainDataGrid);
            this._dataGridGroupBox.Location = new System.Drawing.Point(12, 201);
            this._dataGridGroupBox.Name = "_dataGridGroupBox";
            this._dataGridGroupBox.Size = new System.Drawing.Size(615, 264);
            this._dataGridGroupBox.TabIndex = 4;
            this._dataGridGroupBox.TabStop = false;
            this._dataGridGroupBox.Text = "Key Statistics";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.tbTestInput);
            this.groupBox1.Controls.Add(this.listBoxIntervals);
            this.groupBox1.Controls.Add(this._runAtStartupCheckBox);
            this.groupBox1.Controls.Add(this._startAsAdminCheckBox);
            this.groupBox1.Controls.Add(this._activateOnLaunchCheckBox);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this._minimizeCheckBox);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this._thresholdTimeInput);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(12, 48);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(615, 137);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Settings";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(7, 95);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(66, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "System Start";
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(534, 9);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(73, 26);
            this.label6.TabIndex = 13;
            this.label6.Text = "Diagnose\r\nKeyPress (ms)";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(203, 9);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(294, 39);
            this.label5.TabIndex = 12;
            this.label5.Text = "Type here to get a median value in ms for your typing speed. \r\nSet this value a b" +
    "it lower.\r\ne.g. Typing Speed: 100ms -> Chatter Treshold: 70ms";
            // 
            // tbTestInput
            // 
            this.tbTestInput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbTestInput.Location = new System.Drawing.Point(206, 54);
            this.tbTestInput.Name = "tbTestInput";
            this.tbTestInput.Size = new System.Drawing.Size(326, 78);
            this.tbTestInput.TabIndex = 11;
            this.tbTestInput.Text = "";
            // 
            // listBoxIntervals
            // 
            this.listBoxIntervals.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxIntervals.FormattingEnabled = true;
            this.listBoxIntervals.Location = new System.Drawing.Point(537, 51);
            this.listBoxIntervals.Name = "listBoxIntervals";
            this.listBoxIntervals.Size = new System.Drawing.Size(72, 82);
            this.listBoxIntervals.TabIndex = 10;
            // 
            // _runAtStartupCheckBox
            // 
            this._runAtStartupCheckBox.AutoSize = true;
            this._runAtStartupCheckBox.Location = new System.Drawing.Point(112, 94);
            this._runAtStartupCheckBox.Name = "_runAtStartupCheckBox";
            this._runAtStartupCheckBox.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this._runAtStartupCheckBox.Size = new System.Drawing.Size(15, 14);
            this._runAtStartupCheckBox.TabIndex = 8;
            this._runAtStartupCheckBox.UseVisualStyleBackColor = true;
            this._runAtStartupCheckBox.CheckedChanged += new System.EventHandler(this.OnRunAtStartupCheckBoxChanged);
            // 
            // _startAsAdminCheckBox
            // 
            this._startAsAdminCheckBox.AutoSize = true;
            this._startAsAdminCheckBox.Location = new System.Drawing.Point(112, 116);
            this._startAsAdminCheckBox.Name = "_startAsAdminCheckBox";
            this._startAsAdminCheckBox.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this._startAsAdminCheckBox.Size = new System.Drawing.Size(15, 14);
            this._startAsAdminCheckBox.TabIndex = 15;
            this._startAsAdminCheckBox.UseVisualStyleBackColor = true;
            this._startAsAdminCheckBox.CheckedChanged += new System.EventHandler(this.OnStartAsAdminCheckBoxChanged);
            // 
            // _activateOnLaunchCheckBox
            // 
            this._activateOnLaunchCheckBox.AutoSize = true;
            this._activateOnLaunchCheckBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this._activateOnLaunchCheckBox.Location = new System.Drawing.Point(112, 26);
            this._activateOnLaunchCheckBox.Name = "_activateOnLaunchCheckBox";
            this._activateOnLaunchCheckBox.Size = new System.Drawing.Size(15, 14);
            this._activateOnLaunchCheckBox.TabIndex = 7;
            this._activateOnLaunchCheckBox.TabStop = false;
            this._activateOnLaunchCheckBox.UseVisualStyleBackColor = true;
            this._activateOnLaunchCheckBox.CheckedChanged += new System.EventHandler(this.OnActivateOnLaunchCheckBox);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(7, 116);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(75, 13);
            this.label8.TabIndex = 16;
            this.label8.Text = "Start as Admin";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 26);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Activate on Launch";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(165, 75);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "ms";
            // 
            // _minimizeCheckBox
            // 
            this._minimizeCheckBox.AutoSize = true;
            this._minimizeCheckBox.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this._minimizeCheckBox.Location = new System.Drawing.Point(112, 48);
            this._minimizeCheckBox.Name = "_minimizeCheckBox";
            this._minimizeCheckBox.Size = new System.Drawing.Size(15, 14);
            this._minimizeCheckBox.TabIndex = 4;
            this._minimizeCheckBox.TabStop = false;
            this._minimizeCheckBox.UseVisualStyleBackColor = true;
            this._minimizeCheckBox.CheckedChanged += new System.EventHandler(this.OnMinimizeCheckBoxValueChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(6, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 3;
            this.label3.Text = "Launch Minimized";
            // 
            // _thresholdTimeInput
            // 
            this._thresholdTimeInput.Location = new System.Drawing.Point(112, 68);
            this._thresholdTimeInput.Name = "_thresholdTimeInput";
            this._thresholdTimeInput.Size = new System.Drawing.Size(47, 20);
            this._thresholdTimeInput.TabIndex = 2;
            this._thresholdTimeInput.TabStop = false;
            this._thresholdTimeInput.Value = new decimal(new int[] {
            50,
            0,
            0,
            0});
            this._thresholdTimeInput.ValueChanged += new System.EventHandler(this.OnThresholdTimeInputValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 70);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Chatter Threshold";
            // 
            // _notifyIcon
            // 
            this._notifyIcon.ContextMenuStrip = this._notifyIconMenuStrip;
            this._notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("_notifyIcon.Icon")));
            this._notifyIcon.Text = "Keyboard Unchatter";
            this._notifyIcon.Visible = true;
            this._notifyIcon.MouseClick += new System.Windows.Forms.MouseEventHandler(this.OnNotifyIconClick);
            // 
            // _notifyIconMenuStrip
            // 
            this._notifyIconMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this._openMenuItem,
            this.toolStripMenuItem1,
            this._exitMenuItem});
            this._notifyIconMenuStrip.Name = "_notifyIconMenuStrip";
            this._notifyIconMenuStrip.Size = new System.Drawing.Size(104, 54);
            this._notifyIconMenuStrip.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.OnNotifyIconMenuStripItemClicked);
            // 
            // _openMenuItem
            // 
            this._openMenuItem.Name = "_openMenuItem";
            this._openMenuItem.Size = new System.Drawing.Size(103, 22);
            this._openMenuItem.Text = "Open";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(100, 6);
            // 
            // _exitMenuItem
            // 
            this._exitMenuItem.Name = "_exitMenuItem";
            this._exitMenuItem.Size = new System.Drawing.Size(103, 22);
            this._exitMenuItem.Text = "Exit";
            // 
            // btnReset
            // 
            this.btnReset.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnReset.Location = new System.Drawing.Point(19, 471);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(147, 27);
            this.btnReset.TabIndex = 6;
            this.btnReset.TabStop = false;
            this.btnReset.Text = "Reset Statistics/Diagnose";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.OnResetDiagnosticsClick);
            // 
            // btnRelaunchAsAdmin
            // 
            this.btnRelaunchAsAdmin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRelaunchAsAdmin.Location = new System.Drawing.Point(472, 471);
            this.btnRelaunchAsAdmin.Name = "btnRelaunchAsAdmin";
            this.btnRelaunchAsAdmin.Size = new System.Drawing.Size(155, 27);
            this.btnRelaunchAsAdmin.TabIndex = 7;
            this.btnRelaunchAsAdmin.TabStop = false;
            this.btnRelaunchAsAdmin.Text = "Relaunch as Admin";
            this.btnRelaunchAsAdmin.UseVisualStyleBackColor = true;
            this.btnRelaunchAsAdmin.Click += new System.EventHandler(this.OnRelaunchAsAdminClick);
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 510);
            this.Controls.Add(this.btnRelaunchAsAdmin);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this._dataGridGroupBox);
            this.Controls.Add(this._buttonActivate);
            this.Controls.Add(this._statusPanel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(320, 240);
            this.Name = "MainWindow";
            this.Text = "Keyboard Unchatter";
            this.Resize += new System.EventHandler(this.OnMainWindowResize);
            this._statusPanel.ResumeLayout(false);
            this._statusPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this._mainDataGrid)).EndInit();
            this._dataGridGroupBox.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this._thresholdTimeInput)).EndInit();
            this._notifyIconMenuStrip.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel _statusPanel;
        private System.Windows.Forms.Button _buttonActivate;
        private System.Windows.Forms.DataGridView _mainDataGrid;
        private System.Windows.Forms.GroupBox _dataGridGroupBox;
        private System.Windows.Forms.Label _statusPanelLabel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox _minimizeCheckBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown _thresholdTimeInput;
        private System.Windows.Forms.NotifyIcon _notifyIcon;
        private System.Windows.Forms.ContextMenuStrip _notifyIconMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem _openMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem _exitMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Key;
        private System.Windows.Forms.DataGridViewTextBoxColumn PressCount;
        private System.Windows.Forms.DataGridViewTextBoxColumn ChatterCount;
        private System.Windows.Forms.DataGridViewTextBoxColumn FailureRate;
        private System.Windows.Forms.CheckBox _activateOnLaunchCheckBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox _runAtStartupCheckBox;
        private System.Windows.Forms.CheckBox _startAsAdminCheckBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label labelTypingSpeed;
        private System.Windows.Forms.RichTextBox tbTestInput;
        private System.Windows.Forms.ListBox listBoxIntervals;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnRelaunchAsAdmin;
    }
}

