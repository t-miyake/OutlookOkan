namespace OutlookOkan
{
    partial class ConfirmationWindow
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
            this.sendButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.ToAddressList = new OutlookOkan.CustomCheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.ToLabel = new System.Windows.Forms.Label();
            this.CcAddressList = new OutlookOkan.CustomCheckedListBox();
            this.CcLabel = new System.Windows.Forms.Label();
            this.BccAddressList = new OutlookOkan.CustomCheckedListBox();
            this.BccLabel = new System.Windows.Forms.Label();
            this.AlertBox = new OutlookOkan.CustomCheckedListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.OtherInfoTextBox = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.SubjectTextBox = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.AttachmentsList = new OutlookOkan.CustomCheckedListBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // sendButton
            // 
            this.sendButton.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.sendButton.Enabled = false;
            this.sendButton.Location = new System.Drawing.Point(1062, 573);
            this.sendButton.Margin = new System.Windows.Forms.Padding(2, 3, 9, 4);
            this.sendButton.Name = "sendButton";
            this.sendButton.Size = new System.Drawing.Size(93, 38);
            this.sendButton.TabIndex = 19;
            this.sendButton.Text = "送信";
            this.sendButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(1170, 573);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(2, 3, 9, 4);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(93, 38);
            this.cancelButton.TabIndex = 20;
            this.cancelButton.Text = "キャンセル";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // ToAddressList
            // 
            this.ToAddressList.FormattingEnabled = true;
            this.ToAddressList.Location = new System.Drawing.Point(11, 51);
            this.ToAddressList.Margin = new System.Windows.Forms.Padding(8, 3, 8, 8);
            this.ToAddressList.Name = "ToAddressList";
            this.ToAddressList.ScrollAlwaysVisible = true;
            this.ToAddressList.Size = new System.Drawing.Size(525, 70);
            this.ToAddressList.TabIndex = 7;
            this.ToAddressList.SelectedIndexChanged += new System.EventHandler(this.ToAddressList_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(20, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(214, 22);
            this.label1.TabIndex = 1;
            this.label1.Text = "本当にメールを送信しますか？";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(20, 42);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(415, 19);
            this.label2.TabIndex = 2;
            this.label2.Text = "すべての項目を確認し、チェックを入れて、送信ボタンを押してください。";
            // 
            // ToLabel
            // 
            this.ToLabel.AutoSize = true;
            this.ToLabel.Location = new System.Drawing.Point(9, 27);
            this.ToLabel.Margin = new System.Windows.Forms.Padding(8, 3, 8, 2);
            this.ToLabel.Name = "ToLabel";
            this.ToLabel.Size = new System.Drawing.Size(26, 19);
            this.ToLabel.TabIndex = 6;
            this.ToLabel.Text = "To";
            // 
            // CcAddressList
            // 
            this.CcAddressList.FormattingEnabled = true;
            this.CcAddressList.Location = new System.Drawing.Point(11, 149);
            this.CcAddressList.Margin = new System.Windows.Forms.Padding(8, 3, 8, 8);
            this.CcAddressList.Name = "CcAddressList";
            this.CcAddressList.ScrollAlwaysVisible = true;
            this.CcAddressList.Size = new System.Drawing.Size(525, 70);
            this.CcAddressList.TabIndex = 9;
            this.CcAddressList.SelectedIndexChanged += new System.EventHandler(this.CcAddressList_SelectedIndexChanged_1);
            // 
            // CcLabel
            // 
            this.CcLabel.AutoSize = true;
            this.CcLabel.Location = new System.Drawing.Point(9, 126);
            this.CcLabel.Margin = new System.Windows.Forms.Padding(8, 3, 8, 2);
            this.CcLabel.Name = "CcLabel";
            this.CcLabel.Size = new System.Drawing.Size(29, 19);
            this.CcLabel.TabIndex = 8;
            this.CcLabel.Text = "CC";
            // 
            // BccAddressList
            // 
            this.BccAddressList.FormattingEnabled = true;
            this.BccAddressList.Location = new System.Drawing.Point(11, 248);
            this.BccAddressList.Margin = new System.Windows.Forms.Padding(8, 3, 8, 10);
            this.BccAddressList.Name = "BccAddressList";
            this.BccAddressList.ScrollAlwaysVisible = true;
            this.BccAddressList.Size = new System.Drawing.Size(525, 70);
            this.BccAddressList.TabIndex = 11;
            this.BccAddressList.SelectedIndexChanged += new System.EventHandler(this.BccAddressList_SelectedIndexChanged);
            // 
            // BccLabel
            // 
            this.BccLabel.AutoSize = true;
            this.BccLabel.Location = new System.Drawing.Point(9, 224);
            this.BccLabel.Margin = new System.Windows.Forms.Padding(8, 3, 8, 2);
            this.BccLabel.Name = "BccLabel";
            this.BccLabel.Size = new System.Drawing.Size(39, 19);
            this.BccLabel.TabIndex = 10;
            this.BccLabel.Text = "BCC";
            // 
            // AlertBox
            // 
            this.AlertBox.BackColor = System.Drawing.SystemColors.Window;
            this.AlertBox.Font = new System.Drawing.Font("Meiryo UI", 10.2F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.AlertBox.FormattingEnabled = true;
            this.AlertBox.Location = new System.Drawing.Point(11, 33);
            this.AlertBox.Margin = new System.Windows.Forms.Padding(8, 3, 8, 8);
            this.AlertBox.Name = "AlertBox";
            this.AlertBox.ScrollAlwaysVisible = true;
            this.AlertBox.Size = new System.Drawing.Size(1208, 76);
            this.AlertBox.TabIndex = 4;
            this.AlertBox.SelectedIndexChanged += new System.EventHandler(this.AlertBox_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.BccAddressList);
            this.groupBox1.Controls.Add(this.BccLabel);
            this.groupBox1.Controls.Add(this.CcAddressList);
            this.groupBox1.Controls.Add(this.ToLabel);
            this.groupBox1.Controls.Add(this.CcLabel);
            this.groupBox1.Controls.Add(this.ToAddressList);
            this.groupBox1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(33, 235);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(3, 3, 9, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(549, 332);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "送信先アドレス";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.AlertBox);
            this.groupBox2.Location = new System.Drawing.Point(33, 84);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(1230, 135);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "重要な警告 (念のためメールを再確認してください。)";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.OtherInfoTextBox);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.SubjectTextBox);
            this.groupBox3.Location = new System.Drawing.Point(594, 235);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(3, 3, 9, 4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(669, 115);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "メール情報";
            // 
            // OtherInfoTextBox
            // 
            this.OtherInfoTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.OtherInfoTextBox.Location = new System.Drawing.Point(64, 73);
            this.OtherInfoTextBox.Name = "OtherInfoTextBox";
            this.OtherInfoTextBox.ReadOnly = true;
            this.OtherInfoTextBox.Size = new System.Drawing.Size(594, 27);
            this.OtherInfoTextBox.TabIndex = 16;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 76);
            this.label7.Margin = new System.Windows.Forms.Padding(8, 3, 4, 4);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(47, 19);
            this.label7.TabIndex = 15;
            this.label7.Text = "その他";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(11, 34);
            this.label6.Margin = new System.Windows.Forms.Padding(8, 3, 4, 4);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 19);
            this.label6.TabIndex = 13;
            this.label6.Text = "件名";
            // 
            // SubjectTextBox
            // 
            this.SubjectTextBox.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.SubjectTextBox.Location = new System.Drawing.Point(64, 31);
            this.SubjectTextBox.Name = "SubjectTextBox";
            this.SubjectTextBox.ReadOnly = true;
            this.SubjectTextBox.Size = new System.Drawing.Size(594, 27);
            this.SubjectTextBox.TabIndex = 14;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.AttachmentsList);
            this.groupBox4.Location = new System.Drawing.Point(594, 361);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(3, 3, 9, 4);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(669, 206);
            this.groupBox4.TabIndex = 17;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "添付ファイル";
            // 
            // AttachmentsList
            // 
            this.AttachmentsList.FormattingEnabled = true;
            this.AttachmentsList.Location = new System.Drawing.Point(11, 34);
            this.AttachmentsList.Margin = new System.Windows.Forms.Padding(8, 3, 8, 8);
            this.AttachmentsList.Name = "AttachmentsList";
            this.AttachmentsList.ScrollAlwaysVisible = true;
            this.AttachmentsList.Size = new System.Drawing.Size(647, 158);
            this.AttachmentsList.TabIndex = 18;
            this.AttachmentsList.SelectedIndexChanged += new System.EventHandler(this.AttachmentsList_SelectedIndexChanged);
            // 
            // ConfirmationWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1281, 624);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.sendButton);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConfirmationWindow";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "メール送信前の確認";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button sendButton;
        private System.Windows.Forms.Button cancelButton;
        private CustomCheckedListBox ToAddressList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label ToLabel;
        private CustomCheckedListBox CcAddressList;
        private System.Windows.Forms.Label CcLabel;
        private CustomCheckedListBox BccAddressList;
        private System.Windows.Forms.Label BccLabel;
        private CustomCheckedListBox AlertBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox SubjectTextBox;
        private System.Windows.Forms.GroupBox groupBox4;
        private CustomCheckedListBox AttachmentsList;
        private System.Windows.Forms.TextBox OtherInfoTextBox;
        private System.Windows.Forms.Label label7;
    }
}