namespace OutlookAddIn
{
    partial class SettingWindow
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
            this.tabControl = new System.Windows.Forms.TabControl();
            this.NameAndDomains = new System.Windows.Forms.TabPage();
            this.NameAndDomainsGrid = new System.Windows.Forms.DataGridView();
            this.OkButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.ApplyButton = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CsvImportButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.CsvExportButton = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.WhiteList = new System.Windows.Forms.TabPage();
            this.AlertKeywords = new System.Windows.Forms.TabPage();
            this.AlertAddress = new System.Windows.Forms.TabPage();
            this.AutoCcBccKeywords = new System.Windows.Forms.TabPage();
            this.AutoCcBccAddress = new System.Windows.Forms.TabPage();
            this.tabControl.SuspendLayout();
            this.NameAndDomains.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.NameAndDomainsGrid)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.WhiteList);
            this.tabControl.Controls.Add(this.NameAndDomains);
            this.tabControl.Controls.Add(this.AlertKeywords);
            this.tabControl.Controls.Add(this.AlertAddress);
            this.tabControl.Controls.Add(this.AutoCcBccKeywords);
            this.tabControl.Controls.Add(this.AutoCcBccAddress);
            this.tabControl.Location = new System.Drawing.Point(14, 8);
            this.tabControl.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(924, 371);
            this.tabControl.TabIndex = 0;
            // 
            // NameAndDomains
            // 
            this.NameAndDomains.Controls.Add(this.groupBox3);
            this.NameAndDomains.Controls.Add(this.groupBox2);
            this.NameAndDomains.Controls.Add(this.groupBox1);
            this.NameAndDomains.Location = new System.Drawing.Point(4, 32);
            this.NameAndDomains.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.NameAndDomains.Name = "NameAndDomains";
            this.NameAndDomains.Padding = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.NameAndDomains.Size = new System.Drawing.Size(916, 335);
            this.NameAndDomains.TabIndex = 0;
            this.NameAndDomains.Text = "名称 / ドメイン";
            this.NameAndDomains.UseVisualStyleBackColor = true;
            // 
            // NameAndDomainsGrid
            // 
            this.NameAndDomainsGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.NameAndDomainsGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.NameAndDomainsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.NameAndDomainsGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.NameAndDomainsGrid.Location = new System.Drawing.Point(3, 26);
            this.NameAndDomainsGrid.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.NameAndDomainsGrid.Name = "NameAndDomainsGrid";
            this.NameAndDomainsGrid.RowTemplate.DefaultCellStyle.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.NameAndDomainsGrid.RowTemplate.Height = 24;
            this.NameAndDomainsGrid.Size = new System.Drawing.Size(657, 292);
            this.NameAndDomainsGrid.TabIndex = 0;
            // 
            // OkButton
            // 
            this.OkButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OkButton.Location = new System.Drawing.Point(619, 386);
            this.OkButton.Margin = new System.Windows.Forms.Padding(9, 4, 9, 4);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(93, 38);
            this.OkButton.TabIndex = 1;
            this.OkButton.Text = "OK";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.Location = new System.Drawing.Point(730, 386);
            this.CancelButton.Margin = new System.Windows.Forms.Padding(9, 4, 9, 4);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(93, 38);
            this.CancelButton.TabIndex = 2;
            this.CancelButton.Text = "キャンセル";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ApplyButton
            // 
            this.ApplyButton.Location = new System.Drawing.Point(842, 386);
            this.ApplyButton.Margin = new System.Windows.Forms.Padding(9, 4, 9, 4);
            this.ApplyButton.Name = "ApplyButton";
            this.ApplyButton.Size = new System.Drawing.Size(93, 38);
            this.ApplyButton.TabIndex = 3;
            this.ApplyButton.Text = "適用";
            this.ApplyButton.UseVisualStyleBackColor = true;
            this.ApplyButton.Click += new System.EventHandler(this.ApplyButton_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.NameAndDomainsGrid);
            this.groupBox1.Location = new System.Drawing.Point(6, 7);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(663, 321);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "名称 / ドメイン";
            // 
            // CsvImportButton
            // 
            this.CsvImportButton.Location = new System.Drawing.Point(34, 37);
            this.CsvImportButton.Name = "CsvImportButton";
            this.CsvImportButton.Size = new System.Drawing.Size(166, 38);
            this.CsvImportButton.TabIndex = 2;
            this.CsvImportButton.Text = "CSVインポート";
            this.CsvImportButton.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.CsvExportButton);
            this.groupBox2.Controls.Add(this.CsvImportButton);
            this.groupBox2.Location = new System.Drawing.Point(675, 183);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(8);
            this.groupBox2.Size = new System.Drawing.Size(235, 145);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "インポート/エクスポート";
            // 
            // CsvExportButton
            // 
            this.CsvExportButton.Location = new System.Drawing.Point(34, 89);
            this.CsvExportButton.Margin = new System.Windows.Forms.Padding(3, 8, 3, 3);
            this.CsvExportButton.Name = "CsvExportButton";
            this.CsvExportButton.Size = new System.Drawing.Size(166, 38);
            this.CsvExportButton.TabIndex = 3;
            this.CsvExportButton.Text = "CSVエクスポート";
            this.CsvExportButton.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(675, 7);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(235, 170);
            this.groupBox3.TabIndex = 4;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "設定例";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 8F);
            this.label1.Location = new System.Drawing.Point(8, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(206, 120);
            this.label1.TabIndex = 0;
            this.label1.Text = "名称：のらねこ\r\nドメイン：@noraneko.co.jp\r\n\r\n上記のように、\r\n名称(社名等)とドメインを\r\n登録します。";
            // 
            // WhiteList
            // 
            this.WhiteList.Location = new System.Drawing.Point(4, 32);
            this.WhiteList.Name = "WhiteList";
            this.WhiteList.Size = new System.Drawing.Size(916, 335);
            this.WhiteList.TabIndex = 1;
            this.WhiteList.Text = "ホワイトリスト";
            this.WhiteList.UseVisualStyleBackColor = true;
            // 
            // AlertKeywords
            // 
            this.AlertKeywords.Location = new System.Drawing.Point(4, 32);
            this.AlertKeywords.Name = "AlertKeywords";
            this.AlertKeywords.Size = new System.Drawing.Size(890, 335);
            this.AlertKeywords.TabIndex = 2;
            this.AlertKeywords.Text = "警告キーワード";
            this.AlertKeywords.UseVisualStyleBackColor = true;
            // 
            // AlertAddress
            // 
            this.AlertAddress.Location = new System.Drawing.Point(4, 32);
            this.AlertAddress.Name = "AlertAddress";
            this.AlertAddress.Size = new System.Drawing.Size(890, 335);
            this.AlertAddress.TabIndex = 3;
            this.AlertAddress.Text = "警告アドレス";
            this.AlertAddress.UseVisualStyleBackColor = true;
            // 
            // AutoCcBccKeywords
            // 
            this.AutoCcBccKeywords.Location = new System.Drawing.Point(4, 32);
            this.AutoCcBccKeywords.Name = "AutoCcBccKeywords";
            this.AutoCcBccKeywords.Size = new System.Drawing.Size(890, 335);
            this.AutoCcBccKeywords.TabIndex = 4;
            this.AutoCcBccKeywords.Text = "自動CC/BCC追加(キーワード)";
            this.AutoCcBccKeywords.UseVisualStyleBackColor = true;
            // 
            // AutoCcBccAddress
            // 
            this.AutoCcBccAddress.Location = new System.Drawing.Point(4, 32);
            this.AutoCcBccAddress.Name = "AutoCcBccAddress";
            this.AutoCcBccAddress.Size = new System.Drawing.Size(890, 335);
            this.AutoCcBccAddress.TabIndex = 5;
            this.AutoCcBccAddress.Text = "自動CC/BCC追加(送信先)";
            this.AutoCcBccAddress.UseVisualStyleBackColor = true;
            // 
            // SettingWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 23F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(950, 437);
            this.Controls.Add(this.ApplyButton);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.tabControl);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingWindow";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "設定";
            this.tabControl.ResumeLayout(false);
            this.NameAndDomains.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.NameAndDomainsGrid)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage NameAndDomains;
        private System.Windows.Forms.Button OkButton;
        private new System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Button ApplyButton;
        private System.Windows.Forms.DataGridView NameAndDomainsGrid;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button CsvExportButton;
        private System.Windows.Forms.Button CsvImportButton;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TabPage WhiteList;
        private System.Windows.Forms.TabPage AlertKeywords;
        private System.Windows.Forms.TabPage AlertAddress;
        private System.Windows.Forms.TabPage AutoCcBccKeywords;
        private System.Windows.Forms.TabPage AutoCcBccAddress;
    }
}