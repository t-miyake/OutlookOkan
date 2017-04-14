using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAddIn
{
    public partial class ConfirmWindow : Form
    {

        public ConfirmWindow(Outlook._MailItem mail)
        {
            InitializeComponent();

            GetRecipient(mail);
        }

        public void GetRecipient(Outlook._MailItem mail)
        {
            var displayNameAndRecipient = new Dictionary<string, string>();

            foreach (Outlook.Recipient recip in mail.Recipients)
            {
                var exchangedUser = recip.AddressEntry.GetExchangeUser();
                var registeredUser = recip.AddressEntry.GetContact();

                var nameAndMailAddress = exchangedUser != null
                    ? exchangedUser.Name + @" (" + exchangedUser.PrimarySmtpAddress + @")"
                    : registeredUser != null
                        ? recip.Name
                        : recip.Address;

                displayNameAndRecipient[recip.Name] = nameAndMailAddress;
            }

            var toAdresses = mail.To?.Split(';') ?? new string[] { };
            var ccAdresses = mail.CC?.Split(';') ?? new string[] { };
            var bccAdresses = mail.BCC?.Split(';') ?? new string[] { };

            foreach (var i in displayNameAndRecipient)
            {
                if (toAdresses.Any(address => address.Contains(i.Key)))
                    ToAddressList.Items.Add(i.Value);

                if (ccAdresses.Any(address => address.Contains(i.Key)))
                    CcAddressList.Items.Add(i.Value);

                if (bccAdresses.Any(address => address.Contains(i.Key)))
                    BccAddressList.Items.Add(i.Value);
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