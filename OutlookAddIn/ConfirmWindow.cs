using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace OutlookAddIn
{
    public partial class ConfirmWindow : Form
    {
        public Dictionary<string, string> DisplayNameAndRecipient = new Dictionary<string, string>();

        /// <summary>
        /// メール送信の確認画面を表示。
        /// </summary>
        /// <param name="mail">送信するメールに関する情報</param>
        public ConfirmWindow(Outlook._MailItem mail)
        {
            InitializeComponent();

            DrawRecipient(mail);
            CheckMailbodyAndRecipient(mail);
        }

        /// <summary>
        /// 登録された名称とドメインから、宛先候補ではないアドレスが宛先に含まれている場合に、警告を表示する。
        /// </summary>
        /// <param name="mail"></param>
        public void CheckMailbodyAndRecipient(Outlook._MailItem mail)
        {
            var readCsv = new ReadAndWriteCsv("NameAndDomains.csv");
            var nameAndDomainsList = readCsv.ReadCsv<NameAndDomains>(readCsv.ParseCsv<NameAndDomainsMap>());

            //メールの本文中に、登録された名称があるか確認。
            //var recipientCandidateNames = (from nameAnddomain in nameAndDomainsList where mail.Body.Contains(nameAnddomain.Name) select nameAnddomain.Name).ToList();
            var recipientCandidateDomains = (from nameAnddomain in nameAndDomainsList where mail.Body.Contains(nameAnddomain.Name) select nameAnddomain.Domain).ToList();

            //登録された名称かつ本文中に登場した名称以外のドメインが宛先に含まれている場合、警告を表示。
            //送信先の候補が見つからない場合、何もしない。(見つからない場合の方が多いため、警告ばかりになってしまう。) 
            if (recipientCandidateDomains.Count == 0) return;
            foreach (var recipients in DisplayNameAndRecipient)
            {
                if (recipientCandidateDomains.Any(domains => domains.Equals(recipients.Value.Substring(recipients.Value.IndexOf("@", StringComparison.Ordinal)))))
                {
                    //正常なのでとりあえず何もしない。
                }
                else
                {
                    //送信者ドメインは警告対象外。
                    if (recipients.Value.Contains(mail.SendUsingAccount.SmtpAddress.Substring(mail.SendUsingAccount.SmtpAddress.IndexOf("@", StringComparison.Ordinal)))) continue;
                    AlertBox.Items.Add(recipients.Key + @" : このアドレスは意図した宛先とは無関係の可能性があります！");
                    AlertBox.ColorFlag.Add(true);
                }
            }
        }

        /// <summary>
        /// 送信先メールアドレスを取得し、画面に表示する。
        /// </summary>
        /// <param name="mail">送信するメールに関する情報</param>
        public void DrawRecipient(Outlook._MailItem mail)
        {
            foreach (Outlook.Recipient recip in mail.Recipients)
            {
                // Exchangeの連絡先に登録された情報を取得。
                var exchangeUser = recip.AddressEntry.GetExchangeUser();

                // ローカルの連絡先に登録された情報を取得。
                var registeredUser = recip.AddressEntry.GetContact();

                // 登録されたメールアドレスの場合、登録名のみが表示されるため、メールアドレスと共に表示されるよう表示用テキストを生成。
                var nameAndMailAddress = exchangeUser != null
                    ? exchangeUser.Name + @" (" + exchangeUser.PrimarySmtpAddress + @")"
                    : registeredUser != null
                        ? recip.Name
                        : recip.Address;

                DisplayNameAndRecipient[recip.Name] = nameAndMailAddress;
            }

            // 宛先(To,CC,BCC)に登録された宛先又は登録名を配列に格納。
            var toAdresses = mail.To?.Split(';') ?? new string[] { };
            var ccAdresses = mail.CC?.Split(';') ?? new string[] { };
            var bccAdresses = mail.BCC?.Split(';') ?? new string[] { };

            var senderDomain = mail.SendUsingAccount.SmtpAddress.Substring(mail.SendUsingAccount.SmtpAddress.IndexOf("@", StringComparison.Ordinal));

            // 宛先や登録名から、表示用テキスト(メールアドレスや登録名)を各エリアに表示。
            // 宛先ドメインが送信元ドメインと異なる場合、色を変更するフラグをtrue、そうでない場合falseとする。
            foreach (var i in DisplayNameAndRecipient)
            {
                if (toAdresses.Any(address => address.Contains(i.Key)))
                {
                    ToAddressList.Items.Add(i.Value);
                    ToAddressList.ColorFlag.Add(!i.Value.Contains(senderDomain));
                }

                if (ccAdresses.Any(address => address.Contains(i.Key)))
                {
                    CcAddressList.Items.Add(i.Value);
                    CcAddressList.ColorFlag.Add(!i.Value.Contains(senderDomain));
                }

                if (bccAdresses.Any(address => address.Contains(i.Key)))
                {
                    BccAddressList.Items.Add(i.Value);
                    BccAddressList.ColorFlag.Add(!i.Value.Contains(senderDomain));
                }
            }
        }
        private void AlertBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SendButtonSwitch();
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

        /// <summary>
        /// 全てのチェックボックスにチェックされた場合のみ、送信ボタンを有効とする。
        /// </summary>
        private void SendButtonSwitch()
        {
            if (ToAddressList.CheckedItems.Count == ToAddressList.Items.Count && CcAddressList.CheckedItems.Count == CcAddressList.Items.Count && BccAddressList.CheckedItems.Count == BccAddressList.Items.Count)
            {
                sendButton.Enabled = true;
            }
            else
            {
                sendButton.Enabled = false;
            }
        }
    }
}