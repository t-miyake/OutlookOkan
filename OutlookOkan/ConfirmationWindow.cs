using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace OutlookOkan
{
    // TODO 高速化のため、処理を簡潔にする。
    public partial class ConfirmationWindow : Form
    {
        private readonly Dictionary<string, string> _displayNameAndRecipient = new Dictionary<string, string>();
        private readonly List<Whitelist> _whitelists = new List<Whitelist>();

        /// <summary>
        /// メール送信の確認画面を表示。
        /// </summary>
        /// <param name="mail">送信するメールに関する情報</param>
        public ConfirmationWindow(Outlook._MailItem mail)
        {
            InitializeComponent();

            MakeDisplayNameAndRecipient(mail);

            SubjectTextBox.Text = mail.Subject;
            OtherInfoTextBox.Text = $@"メール種別：{GetMailBodyFormat(mail)}";
            CheckForgotAttach(mail);

            CheckKeyword(mail);
            AutoAddCcAndBcc(mail);

            //TODO 暫定処置だよ！手抜きなだけだよ！
            MakeDisplayNameAndRecipient(mail);
            DrawRecipient(mail);

            DrawAttachments(mail);
            CheckMailbodyAndRecipient(mail);

            UpdateItemsCount();
            SendButtonSwitch();
        }

        /// <summary>
        /// 送信先の表示名と表示名とメールアドレスを対応させるDictionary。(Outlookの仕様上、表示名にメールアドレスが含まれない事がある。)
        /// </summary>
        /// <param name="mail"></param>
        private void MakeDisplayNameAndRecipient(Outlook._MailItem mail)
        {
            foreach (Outlook.Recipient recip in mail.Recipients)
            {
                // Exchangeの連絡先に登録された情報を取得。
                var exchangeUser = recip.AddressEntry.GetExchangeUser();

                // Exchangeの配布リスト(ML)として登録された情報を取得。
                var exchangeDistributionList = recip.AddressEntry.GetExchangeDistributionList();

                // ローカルの連絡先に登録された情報を取得。
                var registeredUser = recip.AddressEntry.GetContact();

                // 登録されたメールアドレスの場合、登録名のみが表示されるため、メールアドレスと共に表示されるよう表示用テキストを生成。
                var nameAndMailAddress = exchangeUser != null
                        ? exchangeUser.Name + $@" ({exchangeUser.PrimarySmtpAddress})" :
                    exchangeDistributionList != null
                        ? exchangeDistributionList.Name + $@" ({exchangeDistributionList.PrimarySmtpAddress})" :
                    registeredUser != null
                        ? recip.Name + $@" ({recip.Address})"
                        : recip.Address;

                _displayNameAndRecipient[recip.Name] = nameAndMailAddress;
            }
        }

        /// <summary>
        /// ファイルの添付忘れを確認。
        /// </summary>
        /// <param name="mail"></param>
        private void CheckForgotAttach(Outlook._MailItem mail)
        {
            if (mail.Body.Contains("添付") && mail.Attachments.Count == 0)
            {
                AlertBox.Items.Add(@"本文中に 添付 という文言があるのに添付ファイルがありません。");
                AlertBox.ColorFlag.Add(false);
            }
        }

        /// <summary>
        /// メールの形式を取得し、表示用の文字列を返す。
        /// </summary>
        /// <param name="mail"></param>
        /// <returns>メールの形式</returns>
        private string GetMailBodyFormat(Outlook._MailItem mail)
        {
            switch (mail.BodyFormat)
            {
                case Outlook.OlBodyFormat.olFormatUnspecified:
                    return "不明";
                case Outlook.OlBodyFormat.olFormatPlain:
                    return "テキスト形式";
                case Outlook.OlBodyFormat.olFormatHTML:
                    return "HTML形式";
                case Outlook.OlBodyFormat.olFormatRichText:
                    return "リッチテキスト形式";
                default:
                    return "不明";
            }
        }

        /// <summary>
        /// 本文に登録したキーワードがある場合、登録した警告文を表示する。
        /// </summary>
        /// <param name="mail"></param>
        private void CheckKeyword(Outlook._MailItem mail)
        {
            //Load AlertKeywordAndMessage
            var readCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            var alertKeywordAndMessageList =
                readCsv.ReadCsv<AlertKeywordAndMessage>(readCsv.ParseCsv<AlertKeywordAndMessageMap>());

            if (alertKeywordAndMessageList.Count != 0)
            {
                foreach (var i in alertKeywordAndMessageList)
                {
                    if (mail.Body.Contains(i.AlertKeyword))
                    {
                        AlertBox.Items.Add(i.Message);
                        AlertBox.ColorFlag.Add(true);
                    }
                }
            }
        }

        private void AutoAddCcAndBcc(Outlook._MailItem mail)
        {
            var autoAddedCcAddressList = new List<string>();
            var autoAddedBccAddressList = new List<string>();

            //Load AutoCcBccKeywordList
            var readCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");
            var autoCcBccKeywordList = readCsv.ReadCsv<AutoCcBccKeyword>(readCsv.ParseCsv<AutoCcBccKeywordMap>());

            if (autoCcBccKeywordList.Count != 0)
            {
                foreach (var i in autoCcBccKeywordList)
                {
                    if (mail.Body.Contains(i.Keyword) && !_displayNameAndRecipient.Any(recip => recip.Key.Contains(i.AutoAddAddress)))
                    {
                        var recip = mail.Recipients.Add(i.AutoAddAddress);
                        recip.Type = i.CcOrBcc == CcOrBcc.CC
                            ? (int)Outlook.OlMailRecipientType.olCC
                            : (int)Outlook.OlMailRecipientType.olBCC;
                        AlertBox.Items.Add($@"自動で {i.CcOrBcc} に {i.AutoAddAddress} が追加されました。(該当キーワード 「{i.Keyword}」)", true);
                        AlertBox.ColorFlag.Add(false);

                        if (i.CcOrBcc == CcOrBcc.CC)
                        {
                            autoAddedCcAddressList.Add(i.AutoAddAddress);
                        }else
                        {
                            autoAddedBccAddressList.Add(i.AutoAddAddress);
                        }

                        // 自動追加されたアドレスはホワイトリスト登録アドレス扱い。
                        _whitelists.Add(new Whitelist { WhiteName = i.AutoAddAddress });
                    }
                }
            }

            //Load AutoCcBccRecipientList
            // TODO 流石にひどいので直す。
            readCsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");
            var autoCcBccRecipientList = readCsv.ReadCsv<AutoCcBccRecipient>(readCsv.ParseCsv<AutoCcBccRecipientMap>());

            if (autoCcBccRecipientList.Count != 0)
            {
                foreach (var i in autoCcBccRecipientList)
                {
                    if (_displayNameAndRecipient.Any(recipient => recipient.Value.Contains(i.TargetRecipient)) && !_displayNameAndRecipient.Any(recip => recip.Key.Contains(i.AutoAddAddress)))
                    {
                        if (i.CcOrBcc == CcOrBcc.CC)
                        {
                            if (!autoAddedCcAddressList.Contains(i.AutoAddAddress))
                            {
                                var recip = mail.Recipients.Add(i.AutoAddAddress);
                                recip.Type = (int)Outlook.OlMailRecipientType.olCC;

                                autoAddedCcAddressList.Add(i.AutoAddAddress);
                            }
                        }else if (!autoAddedBccAddressList.Contains(i.AutoAddAddress))
                            {
                                var recip = mail.Recipients.Add(i.AutoAddAddress);
                                recip.Type = (int)Outlook.OlMailRecipientType.olBCC;

                                autoAddedBccAddressList.Add(i.AutoAddAddress);
                            }
                        
                        AlertBox.Items.Add($@"自動で {i.CcOrBcc} に {i.AutoAddAddress} が追加されました。(該当宛先 「{i.TargetRecipient}」)", true);
                        AlertBox.ColorFlag.Add(false);

                        // 自動追加されたアドレスはホワイトリスト登録アドレス扱い。
                        _whitelists.Add(new Whitelist { WhiteName = i.AutoAddAddress });
                    }
                }
            }

            mail.Recipients.ResolveAll();
        }

        /// <summary>
        /// 添付ファイルとそのファイルサイズを取得し、画面に表示する。
        /// </summary>
        /// <param name="mail"></param>
        private void DrawAttachments(Outlook._MailItem mail)
        {
            if (mail.Attachments.Count != 0)
            {
                for (var i = 0; i < mail.Attachments.Count; i++)
                {
                    AttachmentsList.Items.Add(mail.Attachments[i + 1].FileName + $@" ({(mail.Attachments[i + 1].Size / 1024):N}kB)");
                }
            }
        }

        /// <summary>
        /// 登録された名称とドメインから、宛先候補ではないアドレスが宛先に含まれている場合に、警告を表示する。
        /// </summary>
        /// <param name="mail"></param>
        private void CheckMailbodyAndRecipient(Outlook._MailItem mail)
        {
            //Load NameAndDomainsList
            var readCsv = new ReadAndWriteCsv("NameAndDomains.csv");
            var nameAndDomainsList = readCsv.ReadCsv<NameAndDomains>(readCsv.ParseCsv<NameAndDomainsMap>());
            
            //メールの本文中に、登録された名称があるか確認。
            var recipientCandidateDomains = (from nameAnddomain in nameAndDomainsList where mail.Body.Contains(nameAnddomain.Name) select nameAnddomain.Domain).ToList();

            //登録された名称かつ本文中に登場した名称以外のドメインが宛先に含まれている場合、警告を表示。
            //送信先の候補が見つからない場合、何もしない。(見つからない場合の方が多いため、警告ばかりになってしまう。) 
            if (recipientCandidateDomains.Count != 0)
            {
                foreach (var recipients in _displayNameAndRecipient)
                {
                    if (!recipientCandidateDomains.Any(
                        domains => domains.Equals(
                            recipients.Value.Substring(recipients.Value.IndexOf("@", StringComparison.Ordinal)))))
                    {
                        //送信者ドメインは警告対象外。
                        if (!recipients.Value.Contains(
                            mail.SendUsingAccount.SmtpAddress.Substring(
                                mail.SendUsingAccount.SmtpAddress.IndexOf("@", StringComparison.Ordinal))))
                        {
                            AlertBox.Items.Add(recipients.Key + @" : このアドレスは意図した宛先とは無関係の可能性があります！");
                            AlertBox.ColorFlag.Add(true);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 送信先メールアドレスを取得し、画面に表示する。
        /// </summary>
        /// <param name="mail">送信するメールに関する情報</param>
        private void DrawRecipient(Outlook._MailItem mail)
        {
            // TODO ここでいろいろやりすぎなので、直す。

            //Load Whitelist
            var readCsv = new ReadAndWriteCsv("Whitelist.csv");
            _whitelists.AddRange(readCsv.ReadCsv<Whitelist>(readCsv.ParseCsv<WhitelistMap>()));

            //Load AlertAddressList
            readCsv = new ReadAndWriteCsv("AlertAddressList.csv");
            var alertAddresslist = readCsv.ReadCsv<AlertAddress>(readCsv.ParseCsv<AlertAddressMap>());

            // 宛先(To,CC,BCC)に登録された宛先又は登録名を配列に格納。
            var toAdresses = mail.To?.Split(';') ?? new string[] { };
            var ccAdresses = mail.CC?.Split(';') ?? new string[] { };
            var bccAdresses = mail.BCC?.Split(';') ?? new string[] { };

            var senderDomain = mail.SendUsingAccount.SmtpAddress.Substring(mail.SendUsingAccount.SmtpAddress.IndexOf("@", StringComparison.Ordinal));

            // 宛先や登録名から、表示用テキスト(メールアドレスや登録名)を各エリアに表示。
            // 宛先ドメインが送信元ドメインと異なる場合、色を変更するフラグをtrue、そうでない場合falseとする。
            // ホワイトリストに含まれる宛先の場合、ListのIsCheckedフラグをtrueにして、最初からチェック済みとする。
            // 警告アドレスリストに含まれる宛先の場合、AlertBoxにその旨を追加する。
            foreach (var i in _displayNameAndRecipient)
            {
                if (toAdresses.Any(address => address.Contains(i.Key)))
                {
                    ToAddressList.Items.Add(i.Value, _whitelists.Count != 0 && _whitelists.Any(address => i.Value.Contains(address.WhiteName)));
                    ToAddressList.ColorFlag.Add(!i.Value.Contains(senderDomain));

                    if (alertAddresslist.Count != 0 && alertAddresslist.Any(address => i.Value.Contains(address.TartgetAddress)))
                    {
                        AlertBox.Items.Add($"警告対象として登録されたアドレス/ドメインが宛先(To)に含まれています。 ({i.Value})");
                        AlertBox.ColorFlag.Add(true);
                    }
                }

                if (ccAdresses.Any(address => address.Contains(i.Key)))
                {
                    CcAddressList.Items.Add(i.Value, _whitelists.Count != 0 && _whitelists.Any(address => i.Value.Contains(address.WhiteName)));
                    CcAddressList.ColorFlag.Add(!i.Value.Contains(senderDomain));

                    if (alertAddresslist.Count != 0 && alertAddresslist.Any(address => i.Value.Contains(address.TartgetAddress)))
                    {
                        AlertBox.Items.Add($"警告対象として登録されたアドレス/ドメインが宛先(CC)に含まれています。 ({i.Value})");
                        AlertBox.ColorFlag.Add(true);
                    }
                }

                if (bccAdresses.Any(address => address.Contains(i.Key)))
                {
                    BccAddressList.Items.Add(i.Value, _whitelists.Count != 0 && _whitelists.Any(address => i.Value.Contains(address.WhiteName)));
                    BccAddressList.ColorFlag.Add(!i.Value.Contains(senderDomain));

                    if (alertAddresslist.Count != 0 && alertAddresslist.Any(address => i.Value.Contains(address.TartgetAddress)))
                    {
                        AlertBox.Items.Add($"警告対象として登録されたアドレス/ドメインが宛先(BCC)に含まれています。 ({i.Value})");
                        AlertBox.ColorFlag.Add(true);
                    }
                }
            }
        }

        #region BoxSelectedIndexChanged events
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

        private void AttachmentsList_SelectedIndexChanged(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }
        #endregion

        /// <summary>
        /// 全てのチェックボックスにチェックされた場合のみ、送信ボタンを有効とする。
        /// </summary>
        private void SendButtonSwitch()
        {
            //TODO この判定方法はそのうち直す。
            if (ToAddressList.CheckedItems.Count == ToAddressList.Items.Count && CcAddressList.CheckedItems.Count == CcAddressList.Items.Count && BccAddressList.CheckedItems.Count == BccAddressList.Items.Count && AlertBox.CheckedItems.Count == AlertBox.Items.Count && AttachmentsList.Items.Count == AttachmentsList.CheckedItems.Count)
            {
                sendButton.Enabled = true;
            }
            else
            {
                sendButton.Enabled = false;
            }
        }

        private void UpdateItemsCount()
        {
            ToLabel.Text = $@"To ({ToAddressList.Items.Count})";
            CcLabel.Text = $@"CC ({CcAddressList.Items.Count})";
            BccLabel.Text = $@"BCC ({BccAddressList.Items.Count})";

            AlertAreaGroupBox.Text = $@"重要な警告 ({AlertBox.Items.Count})";
            RecipientGroupBox.Text = $@"送信先アドレス ({ToAddressList.Items.Count + CcAddressList.Items.Count + BccAddressList.Items.Count})";
            AttachmentGroupBox.Text = $@"添付ファイル ({AttachmentsList.Items.Count})";
        }

        // チェックボックスの切り替えを連続して行うと、たまにチェックしていなくても送信ボタンが有効になるので、それを回避。
        // マウスカーソルがエリアから外れた瞬間に再判定する。
        private void AlertBox_MouseLeave(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void ToAddressList_MouseLeave(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void CcAddressList_MouseLeave(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void BccAddressList_MouseLeave(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }

        private void AttachmentsList_MouseLeave(object sender, EventArgs e)
        {
            SendButtonSwitch();
        }
    }
}