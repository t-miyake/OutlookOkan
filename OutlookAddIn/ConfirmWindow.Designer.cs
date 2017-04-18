﻿namespace OutlookAddIn
{
    partial class ConfirmWindow
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
            this.ToAddressList = new OutlookAddIn.CustomCheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.CcAddressList = new OutlookAddIn.CustomCheckedListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.BccAddressList = new OutlookAddIn.CustomCheckedListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.AlertBox = new OutlookAddIn.CustomCheckedListBox();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // sendButton
            // 
            this.sendButton.DialogResult = System.Windows.Forms.DialogResult.Yes;
            this.sendButton.Enabled = false;
            this.sendButton.Location = new System.Drawing.Point(502, 575);
            this.sendButton.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.sendButton.Name = "sendButton";
            this.sendButton.Size = new System.Drawing.Size(93, 38);
            this.sendButton.TabIndex = 0;
            this.sendButton.Text = "送信";
            this.sendButton.UseVisualStyleBackColor = true;
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(616, 575);
            this.cancelButton.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(93, 38);
            this.cancelButton.TabIndex = 1;
            this.cancelButton.Text = "キャンセル";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // ToAddressList
            // 
            this.ToAddressList.FormattingEnabled = true;
            this.ToAddressList.Location = new System.Drawing.Point(32, 285);
            this.ToAddressList.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.ToAddressList.Name = "ToAddressList";
            this.ToAddressList.ScrollAlwaysVisible = true;
            this.ToAddressList.Size = new System.Drawing.Size(651, 48);
            this.ToAddressList.TabIndex = 3;
            this.ToAddressList.SelectedIndexChanged += new System.EventHandler(this.ToAddressList_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(29, 20);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(214, 22);
            this.label1.TabIndex = 4;
            this.label1.Text = "本当にメールを送信しますか？";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(29, 48);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(396, 19);
            this.label2.TabIndex = 5;
            this.label2.Text = "すべてのアドレスを確認し、チェックを入れて、送信を押してください。";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(29, 257);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(26, 19);
            this.label3.TabIndex = 6;
            this.label3.Text = "To";
            // 
            // CcAddressList
            // 
            this.CcAddressList.FormattingEnabled = true;
            this.CcAddressList.Location = new System.Drawing.Point(32, 389);
            this.CcAddressList.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.CcAddressList.Name = "CcAddressList";
            this.CcAddressList.ScrollAlwaysVisible = true;
            this.CcAddressList.Size = new System.Drawing.Size(651, 48);
            this.CcAddressList.TabIndex = 8;
            this.CcAddressList.SelectedIndexChanged += new System.EventHandler(this.CcAddressList_SelectedIndexChanged_1);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(29, 361);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 19);
            this.label4.TabIndex = 9;
            this.label4.Text = "CC";
            // 
            // BccAddressList
            // 
            this.BccAddressList.FormattingEnabled = true;
            this.BccAddressList.Location = new System.Drawing.Point(32, 493);
            this.BccAddressList.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.BccAddressList.Name = "BccAddressList";
            this.BccAddressList.ScrollAlwaysVisible = true;
            this.BccAddressList.Size = new System.Drawing.Size(651, 48);
            this.BccAddressList.TabIndex = 10;
            this.BccAddressList.SelectedIndexChanged += new System.EventHandler(this.BccAddressList_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(29, 465);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(39, 19);
            this.label5.TabIndex = 11;
            this.label5.Text = "BCC";
            // 
            // AlertBox
            // 
            this.AlertBox.BackColor = System.Drawing.SystemColors.Window;
            this.AlertBox.Font = new System.Drawing.Font("Meiryo UI", 10.2F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.AlertBox.FormattingEnabled = true;
            this.AlertBox.HorizontalScrollbar = true;
            this.AlertBox.Location = new System.Drawing.Point(32, 138);
            this.AlertBox.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.AlertBox.Name = "AlertBox";
            this.AlertBox.ScrollAlwaysVisible = true;
            this.AlertBox.Size = new System.Drawing.Size(651, 80);
            this.AlertBox.TabIndex = 12;
            this.AlertBox.SelectedIndexChanged += new System.EventHandler(this.AlertBox_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(29, 109);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(327, 19);
            this.label6.TabIndex = 13;
            this.label6.Text = "重要な警告 (念のため、メールを再確認してください。)";
            // 
            // ConfirmWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(720, 624);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.AlertBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.BccAddressList);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CcAddressList);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ToAddressList);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.sendButton);
            this.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(2, 3, 2, 3);
            this.Name = "ConfirmWindow";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "メール送信前の確認です。";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button sendButton;
        private System.Windows.Forms.Button cancelButton;
        private CustomCheckedListBox ToAddressList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private CustomCheckedListBox CcAddressList;
        private System.Windows.Forms.Label label4;
        private CustomCheckedListBox BccAddressList;
        private System.Windows.Forms.Label label5;
        private CustomCheckedListBox AlertBox;
        private System.Windows.Forms.Label label6;
    }
}