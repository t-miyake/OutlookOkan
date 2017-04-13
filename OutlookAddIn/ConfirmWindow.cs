using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class ConfirmWindow : Form
    {

        public ConfirmWindow(Outlook.MailItem mail)
        {
            InitializeComponent();

            var toAddresses = mail.To.Split(';');
            foreach(var to in toAddresses)
            {
                ToAddressList.Items.Add(to);
            }

            if (mail.CC != null)
            {
                var ccAdresses = mail.CC.Split(';');
                foreach (var cc in ccAdresses)
                {
                    CcAddressList.Items.Add(cc);
                }
            }

            if (mail.BCC != null)
            {
                var bccAdresses = mail.BCC.Split(';');
                foreach (var bcc in bccAdresses)
                {
                    BccAddressList.Items.Add(bcc);
                }
            }
        }

        private void ToAddressList_SelectedIndexChanged(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void CcAddressList_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void BccAddressList_SelectedIndexChanged(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void SendButtonSwitch()
        {
            if (ToAddressList.CheckedItems.Count == ToAddressList.Items.Count && CcAddressList.CheckedItems.Count == CcAddressList.Items.Count && BccAddressList.CheckedItems.Count == BccAddressList.Items.Count)
            {
                sendButton.Enabled = true;
            }else
            {
                sendButton.Enabled = false;
            }
        }

    }
}