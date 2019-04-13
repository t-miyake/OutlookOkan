using OutlookOkan.CsvTools;
using OutlookOkan.Properties;
using OutlookOkan.Types;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Models
{
    public class GenerateCheckList
    {
        private readonly CheckList _checkList = new CheckList();

        private readonly Dictionary<string, string> _displayNameAndRecipient = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _toDisplayNameAndRecipient = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _ccDisplayNameAndRecipient = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _bccDisplayNameAndRecipient = new Dictionary<string, string>();

        private readonly List<Whitelist> _whitelist = new List<Whitelist>();

        /// <summary>
        /// メール送信の確認画面を表示。
        /// </summary>
        /// <param name="mail">送信するメールアイテム</param>
        /// <param name="generalSetting">一般設定</param>
        public CheckList GenerateCheckListFromMail(Outlook._MailItem mail, GeneralSetting generalSetting)
        {
            //This methods must run first.
            GetGeneralMailInfomation(in mail);

            MakeDisplayNameAndRecipient(mail.Recipients, generalSetting);

            CheckForgotAttach(in mail);

            CheckKeyword();

            AutoAddCcAndBcc(mail, generalSetting);

            GetRecipient();

            GetAttachmentsInfomation(in mail);

            CheckMailbodyAndRecipient();

            _checkList.RecipientExternalDomainNum = CountRecipientExternalDomains();

            _checkList.DeferredMinutes = CalcDeferredMinutes();

            return _checkList;
        }

        /// <summary>
        /// 一般的なメールの情報を取得して格納する。
        /// </summary>
        /// <param name="mail">Mail</param>
        private void GetGeneralMailInfomation(in Outlook._MailItem mail)
        {
            if (string.IsNullOrEmpty(mail.SentOnBehalfOfName))
            {
                //mail.SenderEmailAddressがExchangeのアカウントだとそのまま使えないので、メールアドレスを取得する。
                if (mail.SenderEmailType == "EX")
                {
                    var tempOutlookApp = new Outlook.Application();
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(mail.SenderEmailAddress);
                    var exchangeUser = tempRecipient.AddressEntry.GetExchangeUser();

                    _checkList.Sender = exchangeUser.PrimarySmtpAddress ?? Resources.FailedToGetInformation;
                }
                else
                {
                    _checkList.Sender = mail.SenderEmailAddress ?? Resources.FailedToGetInformation;
                }

                if (!_checkList.Sender.Contains("@"))
                {
                    _checkList.Sender = Resources.FailedToGetInformation;
                }

                _checkList.SenderDomain = _checkList.Sender == Resources.FailedToGetInformation ? "------------------" : _checkList.Sender.Substring(_checkList.Sender.IndexOf("@", StringComparison.Ordinal));
            }
            else
            {
                //代理送信の場合
                _checkList.Sender = mail.Sender.Address ?? Resources.FailedToGetInformation;

                if (_checkList.Sender.Contains("@"))
                {
                    //メールアドレスが取得できる場合はExchangeではないのでそのままでよい。
                    _checkList.Sender = $@"{_checkList.Sender} ([{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                    _checkList.SenderDomain = _checkList.Sender.Substring(_checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                }
                else
                {
                    //代理送信の場合かつExchange利用
                    var tempOutlookApp = new Outlook.Application();
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(mail.Sender.Address);

                    _checkList.Sender = $@"[{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf}";
                    _checkList.SenderDomain = @"------------------";

                    if (tempRecipient.AddressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)
                    {
                        //ユーザの代理送信
                        var exchangeUser = tempRecipient.AddressEntry.GetExchangeUser();
                        _checkList.Sender = $@"{exchangeUser.PrimarySmtpAddress} ([{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                        _checkList.SenderDomain = exchangeUser.PrimarySmtpAddress.Substring(exchangeUser.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                    }

                    if (tempRecipient.AddressEntry.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry)
                    {
                        //配布リストの代理送信
                        var exchangeDistributionList = tempRecipient.AddressEntry.GetExchangeDistributionList();
                        _checkList.Sender = $@"{exchangeDistributionList.PrimarySmtpAddress} ([{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                        _checkList.SenderDomain = exchangeDistributionList.PrimarySmtpAddress.Substring(exchangeDistributionList.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                    }
                }
            }

            _checkList.Subject = mail.Subject ?? Resources.FailedToGetInformation;
            _checkList.MailType = GetMailBodyFormat(mail.BodyFormat) ?? Resources.FailedToGetInformation;
            _checkList.MailBody = mail.Body ?? Resources.FailedToGetInformation;

            //改行が2行になる問題を回避するため、HTML形式の場合にのみ2行の改行を1行に置換する。
            if (_checkList.MailType == Resources.HTML)
            {
                _checkList.MailBody = _checkList.MailBody.Replace("\r\n\r\n", "\r\n");
            }

            _checkList.MailHtmlBody = mail.HTMLBody ?? Resources.FailedToGetInformation;
        }

        /// <summary>
        /// 送信者ドメインを除く宛先のドメイン数を数える。
        /// </summary>
        /// <returns>送信者ドメインを除く宛先のドメイン数</returns>
        private int CountRecipientExternalDomains()
        {
            var domainList = new HashSet<string>();
            foreach (var mail in _displayNameAndRecipient)
            {
                var recipient = mail.Key;
                if (recipient != Resources.FailedToGetInformation && recipient.Contains("@"))
                {
                    domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }

            //外部ドメインの数のため、送信者のドメインが含まれていた場合それをマイナスする。
            if (domainList.Contains(_checkList.SenderDomain))
            {
                return domainList.Count - 1;
            }

            return domainList.Count;
        }

        /// <summary>
        /// 宛先メールアドレスと宛先名称を取得する。
        /// </summary>
        /// <param name="recip">メールの宛先</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetNameAndRecipient(Outlook.Recipient recip)
        {
            var mailAddress = Resources.FailedToGetInformation;
            try
            {
                mailAddress =
                    recip.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                        .ToString() ?? Resources.FailedToGetInformation;
            }
            catch (Exception)
            {
                // Do Nothing.
            }

            string nameAndMailAddress;
            if (string.IsNullOrEmpty(recip.Name))
            {
                nameAndMailAddress = mailAddress ?? Resources.FailedToGetInformation;
            }
            else
            {
                nameAndMailAddress = recip.Name.Contains($@" ({mailAddress})") ? recip.Name : recip.Name + $@" ({mailAddress})";
            }

            //ケースによってメールアドレスのみを正しく取得できない恐れがあるため、その場合は、表示名称をメールアドレスとして登録する。
            if (mailAddress?.Contains("@") != true)
            {
                mailAddress = nameAndMailAddress;
            }

            return new List<NameAndRecipient> { new NameAndRecipient { MailAddress = mailAddress, NameAndMailAddress = nameAndMailAddress } };
        }

        /// <summary>
        /// Exchangeの配布リストを展開して宛先メールアドレスと宛先名称を取得する。(入れ子は非展開)
        /// </summary>
        /// <param name="recip">メールの宛先</param>
        /// <param name="enableGetExchangeDistributionListMembers">配布リスト展開のオンオフ設定</param>
        /// <param name="exchangeDistributionListMembersAreWhite">配布リストで展開したアドレスをホワイトリスト化するか否かの設定</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetExchangeDistributionListMembers(Outlook.Recipient recip, bool enableGetExchangeDistributionListMembers, bool exchangeDistributionListMembersAreWhite)
        {
            if (recip.AddressEntry.AddressEntryUserType != Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry) return null;

            try
            {
                var distributionList = recip.AddressEntry.GetExchangeDistributionList();
                var addressEntries = distributionList.GetExchangeDistributionListMembers();

                if (addressEntries is null) return null;

                var exchangeDistributionListMembers = new List<NameAndRecipient>();

                if (addressEntries.Count == 0 || !enableGetExchangeDistributionListMembers)
                {
                    exchangeDistributionListMembers.Add(new NameAndRecipient { MailAddress = distributionList.PrimarySmtpAddress, NameAndMailAddress = distributionList.Name + $@" ({distributionList.PrimarySmtpAddress})" });

                    if (exchangeDistributionListMembersAreWhite)
                    {
                        _whitelist.Add(new Whitelist { WhiteName = distributionList.PrimarySmtpAddress });
                    }

                    return exchangeDistributionListMembers;
                }

                var tempOutlookApp = new Outlook.Application();
                foreach (Outlook.AddressEntry member in addressEntries)
                {
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(member.Address);
                    var mailAddress = Resources.FailedToGetInformation;

                    try
                    {
                        //メールアドレスをExchangeの登録情報から取得。
                        mailAddress = tempRecipient.AddressEntry.PropertyAccessor
                                .GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation;
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }

                    // 入れ子になった配布リストは展開しない。(Exchangeサーバへの負荷が大きく時間もかかるため)
                    exchangeDistributionListMembers.Add(new NameAndRecipient { MailAddress = mailAddress, NameAndMailAddress = (member.Name ?? Resources.FailedToGetInformation) + $@" ({mailAddress})", IncludedGroupAndList = $@" [{distributionList.Name}]" });

                    if (exchangeDistributionListMembersAreWhite)
                    {
                        _whitelist.Add(new Whitelist { WhiteName = mailAddress });
                    }
                }

                return exchangeDistributionListMembers;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 連絡先グループを展開して宛先メールアドレスと宛先名称を取得する。(入れ子も自動展開)
        /// </summary>
        /// <param name="recip">メールの宛先</param>
        /// <param name="contactGroupId">既に確認したGroupID</param>
        /// <param name="enableGetContactGroupMembers">連絡先グループ展開のオンオフ設定</param>
        /// <param name="contactGroupMembersAreWhite">連絡先グループで展開したアドレスをホワイトリスト化するか否かの設定</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetContactGroupMembers(Outlook.Recipient recip, string contactGroupId, bool enableGetContactGroupMembers, bool contactGroupMembersAreWhite)
        {
            var contactGroupMembers = new List<NameAndRecipient>();
            if (!enableGetContactGroupMembers)
            {
                contactGroupMembers.Add(new NameAndRecipient { MailAddress = recip.Name, NameAndMailAddress = recip.Name + $@" [{Resources.ContactGroup}]" });
                return contactGroupMembers;
            }

            string entryId;
            if (contactGroupId is null)
            {
                var entryIdLength = Convert.ToInt32(recip.AddressEntry.ID.Substring(66, 2) + recip.AddressEntry.ID.Substring(64, 2), 16) * 2;
                entryId = recip.AddressEntry.ID.Substring(72, entryIdLength);
            }
            else
            {
                //入れ子の場合のID
                entryId = recip.AddressEntry.ID.Substring(42);
            }

            if (contactGroupId?.Contains(entryId) == true) return null;

            contactGroupId = contactGroupId + entryId + ",";

            var tempOutlookApp = new Outlook.Application().GetNamespace("MAPI");
            var distList = (Outlook.DistListItem)tempOutlookApp.GetItemFromID(entryId);

            for (var i = 1; i < distList.MemberCount + 1; i++)
            {
                var member = distList.GetMember(i);
                contactGroupMembers.AddRange(member.Address == "Unknown"
                    ? GetContactGroupMembers(member, contactGroupId, enableGetContactGroupMembers, contactGroupMembersAreWhite)
                    : GetNameAndRecipient(member));
            }

            foreach (var item in contactGroupMembers)
            {
                item.IncludedGroupAndList += $@" [{distList.DLName}]";

                if (contactGroupMembersAreWhite)
                {
                    _whitelist.Add(new Whitelist { WhiteName = item.MailAddress });
                }
            }

            return contactGroupMembers;
        }

        /// <summary>
        /// 送信先の表示名と表示名とメールアドレスを対応させる。(Outlookの仕様上、表示名にメールアドレスが含まれない事がある。)
        /// </summary>
        /// <param name="recipients">メールの宛先</param>
        /// <param name="generalSetting">一般設定</param>
        private void MakeDisplayNameAndRecipient(IEnumerable recipients, GeneralSetting generalSetting)
        {
            foreach (Outlook.Recipient recip in recipients)
            {
                var nameAndRecipient = new List<NameAndRecipient>();

                switch (recip.AddressEntry.AddressEntryUserType)
                {
                    case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                        nameAndRecipient.AddRange(GetExchangeDistributionListMembers(recip, generalSetting.EnableGetExchangeDistributionListMembers, generalSetting.ExchangeDistributionListMembersAreWhite));
                        break;
                    case Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry:
                        nameAndRecipient.AddRange(GetContactGroupMembers(recip, null, generalSetting.EnableGetContactGroupMembers, generalSetting.ContactGroupMembersAreWhite));
                        break;
                    default:
                        nameAndRecipient.AddRange(GetNameAndRecipient(recip));
                        break;
                }

                foreach (var item in nameAndRecipient)
                {
                    if (_displayNameAndRecipient.ContainsKey(item.MailAddress))
                    {
                        _displayNameAndRecipient[item.MailAddress] += item.IncludedGroupAndList;
                    }
                    else
                    {
                        _displayNameAndRecipient[item.MailAddress] = item.NameAndMailAddress + item.IncludedGroupAndList;
                    }

                    //TODO Temporary processing. It will be improved.
                    //名称と差出人とメールアドレスの紐づけをTo/CC/BCCそれぞれに格納
                    switch (recip.Type)
                    {
                        case 1:
                            if (_toDisplayNameAndRecipient.ContainsKey(item.MailAddress))
                            {
                                _toDisplayNameAndRecipient[item.MailAddress] += item.IncludedGroupAndList;
                            }
                            else
                            {
                                _toDisplayNameAndRecipient[item.MailAddress] = item.NameAndMailAddress + item.IncludedGroupAndList;
                            }
                            continue;
                        case 2:
                            if (_ccDisplayNameAndRecipient.ContainsKey(item.MailAddress))
                            {
                                _ccDisplayNameAndRecipient[item.MailAddress] += item.IncludedGroupAndList;
                            }
                            else
                            {
                                _ccDisplayNameAndRecipient[item.MailAddress] = item.NameAndMailAddress + item.IncludedGroupAndList;
                            }
                            continue;
                        case 3:
                            if (_bccDisplayNameAndRecipient.ContainsKey(item.MailAddress))
                            {
                                _bccDisplayNameAndRecipient[item.MailAddress] += item.IncludedGroupAndList;
                            }
                            else
                            {
                                _bccDisplayNameAndRecipient[item.MailAddress] = item.NameAndMailAddress + item.IncludedGroupAndList;
                            }
                            continue;
                        default:
                            continue;
                    }
                }
            }
        }

        /// <summary>
        /// ファイルの添付忘れを確認。
        /// </summary>
        /// <param name="mail">Mail</param>
        private void CheckForgotAttach(in Outlook._MailItem mail)
        {
            if (mail.Attachments.Count >= 1) return;

            var generalSetting = new List<GeneralSetting>();
            var readCsv = new ReadAndWriteCsv("GeneralSetting.csv");
            foreach (var data in readCsv.GetCsvRecords<GeneralSetting>(readCsv.LoadCsv<GeneralSettingMap>()))
            {
                generalSetting.Add(data);
            }

            string attachmentsKeyword;

            if (generalSetting.Count == 0)
            {
                attachmentsKeyword = "添付";
            }
            else
            {
                if (!generalSetting[0].EnableForgottenToAttachAlert) return;

                switch (generalSetting[0].LanguageCode)
                {
                    case "ja-JP":
                        attachmentsKeyword = "添付";
                        break;
                    case "en-US":
                        attachmentsKeyword = "attached file";
                        break;
                    default:
                        return;
                }
            }

            if (_checkList.MailBody.Contains(attachmentsKeyword))
            {
                _checkList.Alerts.Add(new Alert { AlertMessage = Resources.ForgottenToAttachAlert, IsImportant = true, IsWhite = false, IsChecked = false });
            }
        }

        /// <summary>
        /// メールの形式を取得し、表示用の文字列を返す。
        /// </summary>
        /// <param name="bodyFormat">メールのフォーマット</param>
        /// <returns>メールの形式</returns>
        private string GetMailBodyFormat(Outlook.OlBodyFormat bodyFormat)
        {
            switch (bodyFormat)
            {
                case Outlook.OlBodyFormat.olFormatUnspecified:
                    return Resources.Unknown;
                case Outlook.OlBodyFormat.olFormatPlain:
                    return Resources.Text;
                case Outlook.OlBodyFormat.olFormatHTML:
                    return Resources.HTML;
                case Outlook.OlBodyFormat.olFormatRichText:
                    return Resources.RichText;
                default:
                    return Resources.Unknown;
            }
        }

        /// <summary>
        /// 本文に登録したキーワードがある場合、登録した警告文を表示する。
        /// </summary>
        private void CheckKeyword()
        {
            //Load AlertKeywordAndMessage
            var csv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            var alertKeywordAndMessageList = csv.GetCsvRecords<AlertKeywordAndMessage>(csv.LoadCsv<AlertKeywordAndMessageMap>());

            if (alertKeywordAndMessageList.Count == 0) return;
            foreach (var i in alertKeywordAndMessageList)
            {
                if (!_checkList.MailBody.Contains(i.AlertKeyword)) continue;

                _checkList.Alerts.Add(new Alert { AlertMessage = i.Message, IsImportant = true, IsWhite = false, IsChecked = false });

                if (!i.IsCanNotSend) continue;

                _checkList.IsCanNotSendMail = true;
                _checkList.CanNotSendMailMessage = i.Message;
            }
        }

        /// <summary>
        /// 条件に一致した場合、CCやBCCに宛先を追加する。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <param name="generalSetting">一般設定</param>
        private void AutoAddCcAndBcc(Outlook._MailItem mail, GeneralSetting generalSetting)
        {
            var autoAddedCcAddressList = new List<string>();
            var autoAddedBccAddressList = new List<string>();
            var autoAddRecipients = new List<Outlook.Recipient>();

            //Load AutoCcBccKeywordList
            var autoCcBccKeywordListCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");
            var autoCcBccKeywordList = autoCcBccKeywordListCsv.GetCsvRecords<AutoCcBccKeyword>(autoCcBccKeywordListCsv.LoadCsv<AutoCcBccKeywordMap>());

            if (autoCcBccKeywordList.Count != 0)
            {
                foreach (var i in autoCcBccKeywordList)
                {
                    if (!_checkList.MailBody.Contains(i.Keyword) || !i.AutoAddAddress.Contains("@")) continue;

                    if (i.CcOrBcc == CcOrBcc.CC)
                    {
                        if (!autoAddedCcAddressList.Contains(i.AutoAddAddress) && !_ccDisplayNameAndRecipient.ContainsKey(i.AutoAddAddress))
                        {
                            var recip = mail.Recipients.Add(i.AutoAddAddress);
                            recip.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recip);
                            autoAddedCcAddressList.Add(i.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(i.AutoAddAddress) && !_bccDisplayNameAndRecipient.ContainsKey(i.AutoAddAddress))
                    {
                        var recip = mail.Recipients.Add(i.AutoAddAddress);
                        recip.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recip);
                        autoAddedBccAddressList.Add(i.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{i.CcOrBcc}] [{i.AutoAddAddress}] (" + Resources.ApplicableKeywords + $" 「{i.Keyword}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    // 自動追加されたアドレスはホワイトリスト登録アドレス扱い。
                    _whitelist.Add(new Whitelist { WhiteName = i.AutoAddAddress });
                }
            }

            //Load AutoCcBccRecipientList
            // TODO To be improved
            var autoCcBccRecipientListcsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");
            var autoCcBccRecipientList = autoCcBccRecipientListcsv.GetCsvRecords<AutoCcBccRecipient>(autoCcBccRecipientListcsv.LoadCsv<AutoCcBccRecipientMap>());

            if (autoCcBccRecipientList.Count != 0)
            {
                foreach (var i in autoCcBccRecipientList)
                {
                    if (!_displayNameAndRecipient.Any(recipient => recipient.Key.Contains(i.TargetRecipient)) || !i.AutoAddAddress.Contains("@")) continue;

                    if (i.CcOrBcc == CcOrBcc.CC)
                    {
                        if (!autoAddedCcAddressList.Contains(i.AutoAddAddress) && !_ccDisplayNameAndRecipient.ContainsKey(i.AutoAddAddress))
                        {
                            var recip = mail.Recipients.Add(i.AutoAddAddress);
                            recip.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recip);
                            autoAddedCcAddressList.Add(i.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(i.AutoAddAddress) && !_bccDisplayNameAndRecipient.ContainsKey(i.AutoAddAddress))
                    {
                        var recip = mail.Recipients.Add(i.AutoAddAddress);
                        recip.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recip);
                        autoAddedBccAddressList.Add(i.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{i.CcOrBcc}] [{i.AutoAddAddress}] (" + Resources.ApplicableDestination + $" 「{i.TargetRecipient}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    // 自動追加されたアドレスはホワイトリスト登録アドレス扱い。
                    _whitelist.Add(new Whitelist { WhiteName = i.AutoAddAddress });
                }
            }

            if (autoAddRecipients.Count != 0)
            {
                MakeDisplayNameAndRecipient(autoAddRecipients, generalSetting);
            }

            mail.Recipients.ResolveAll();
        }


        /// <summary>
        /// 添付ファイルとそのファイルサイズを取得し、チェックリストに追加する。
        /// </summary>
        /// <param name="mail"></param>
        private void GetAttachmentsInfomation(in Outlook._MailItem mail)
        {
            if (mail.Attachments.Count == 0) return;

            for (var i = 0; i < mail.Attachments.Count; i++)
            {
                var fileSize = Math.Round(((double)mail.Attachments[i + 1].Size / 1024), 0, MidpointRounding.AwayFromZero).ToString("##,###") + "KB";

                //10Mbyte以上の添付ファイルは警告も表示。
                if (mail.Attachments[i + 1].Size >= 10485760)
                {
                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsBigAttachedFile + $"[{mail.Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                }

                //一部の状態で添付ファイルのファイルタイプを取得できないため、それを回避。
                string fileType;
                try
                {
                    fileType = mail.Attachments[i + 1].FileName.Substring(mail.Attachments[i + 1].FileName.LastIndexOf(".", StringComparison.Ordinal));
                }
                catch (Exception)
                {
                    fileType = Resources.Unknown;
                }

                var isDangerous = false;

                //実行ファイル(.exe)を添付していたら警告を表示
                if (fileType == ".exe")
                {
                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsAttachedExe + $"[{mail.Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                    isDangerous = true;
                }

                string attachmetName;
                try
                {
                    attachmetName = mail.Attachments[i + 1].FileName;
                }
                catch (Exception)
                {
                    attachmetName = Resources.Unknown;
                }

                _checkList.Attachments.Add(new Attachment { FileName = attachmetName, FileSize = fileSize, FileType = fileType, IsTooBig = mail.Attachments[i + 1].Size >= 10485760, IsEncrypted = false, IsChecked = false, IsDangerous = isDangerous });
            }
        }

        /// <summary>
        /// 登録された名称とドメインから、宛先候補ではないアドレスが宛先に含まれている場合に、警告を表示する。
        /// </summary>
        private void CheckMailbodyAndRecipient()
        {
            //Load NameAndDomainsList
            var csv = new ReadAndWriteCsv("NameAndDomains.csv");
            var nameAndDomainsList = csv.GetCsvRecords<NameAndDomains>(csv.LoadCsv<NameAndDomainsMap>());

            //メールの本文中に、登録された名称があるか確認。
            var recipientCandidateDomains = (from nameAnddomain in nameAndDomainsList where _checkList.MailBody.Contains(nameAnddomain.Name) select nameAnddomain.Domain).ToList();

            //登録された名称かつ本文中に登場した名称以外のドメインが宛先に含まれている場合、警告を表示。
            //送信先の候補が見つからない場合、何もしない。(見つからない場合の方が多いため、警告ばかりになってしまう。)
            if (recipientCandidateDomains.Count == 0) return;

            foreach (var recipients in _displayNameAndRecipient)
            {
                if (recipientCandidateDomains.Any(domains => domains.Equals(recipients.Key.Substring(recipients.Key.IndexOf("@", StringComparison.Ordinal))))) continue;

                //送信者ドメインは警告対象外。
                if (!recipients.Key.Contains(_checkList.SenderDomain))
                {
                    _checkList.Alerts.Add(new Alert { AlertMessage = recipients.Value + " : " + Resources.IsAlertAddressMaybeIrrelevant, IsImportant = true, IsWhite = false, IsChecked = false });
                }
            }
        }

        /// <summary>
        /// 送信先メールアドレスを取得し、チェックリストに追加する。
        /// </summary>
        private void GetRecipient()
        {
            // TODO To be improved

            //Load Whitelist
            var readCsv = new ReadAndWriteCsv("Whitelist.csv");
            _whitelist.AddRange(readCsv.GetCsvRecords<Whitelist>(readCsv.LoadCsv<WhitelistMap>()));

            //Load AlertAddressList
            readCsv = new ReadAndWriteCsv("AlertAddressList.csv");
            var alertAddresslist = readCsv.GetCsvRecords<AlertAddress>(readCsv.LoadCsv<AlertAddressMap>());

            // 宛先や登録名から、表示用テキスト(メールアドレスや登録名)を各エリアに表示。
            // 宛先ドメインが送信元ドメインと異なる場合、色を変更するフラグをtrue、そうでない場合falseとする。
            // ホワイトリストに含まれる宛先の場合、ListのIsCheckedフラグをtrueにして、最初からチェック済みとする。
            // 警告アドレスリストに含まれる宛先の場合、AlertBoxにその旨を追加する。

            //TODO 重複が多いので切り出してまとめる。
            foreach (var i in _toDisplayNameAndRecipient)
            {
                var isExternal = !i.Key.Contains(_checkList.SenderDomain);
                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => i.Key.Contains(x.WhiteName));
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist)
                    {
                        if (i.Key.Contains(whitelist.WhiteName))
                        {
                            isSkip = whitelist.IsSkipConfirmation;
                        }
                    }
                }

                _checkList.ToAddresses.Add(new Address { MailAddress = i.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddresslist.Count == 0 || !alertAddresslist.Any(address => i.Key.Contains(address.TartgetAddress))) continue;

                _checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.IsAlertAddressToAlert + $"[{i.Value}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                //送信禁止アドレスに該当する場合、禁止フラグを立て対象メールアドレスを説明文へ追加。
                foreach (var alertAddress in alertAddresslist)
                {
                    if (!i.Key.Contains(alertAddress.TartgetAddress) || !alertAddress.IsCanNotSend) continue;

                    _checkList.IsCanNotSendMail = true;
                    _checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{i.Value}]";
                }
            }

            foreach (var i in _ccDisplayNameAndRecipient)
            {
                var isExternal = !i.Key.Contains(_checkList.SenderDomain);
                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => i.Key.Contains(x.WhiteName));
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist)
                    {
                        if (i.Key.Contains(whitelist.WhiteName))
                        {
                            isSkip = whitelist.IsSkipConfirmation;
                        }
                    }
                }

                _checkList.CcAddresses.Add(new Address { MailAddress = i.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddresslist.Count == 0 || !alertAddresslist.Any(address => i.Key.Contains(address.TartgetAddress))) continue;

                _checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.IsAlertAddressCcAlert + $"[{i.Value}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                //送信禁止アドレスに該当する場合、禁止フラグを立て対象メールアドレスを説明文へ追加。
                foreach (var alertAddress in alertAddresslist)
                {
                    if (!i.Key.Contains(alertAddress.TartgetAddress) || !alertAddress.IsCanNotSend) continue;

                    _checkList.IsCanNotSendMail = true;
                    _checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{i.Value}]";
                }
            }

            foreach (var i in _bccDisplayNameAndRecipient)
            {
                var isExternal = !i.Key.Contains(_checkList.SenderDomain);
                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => i.Key.Contains(x.WhiteName));
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist)
                    {
                        if (i.Key.Contains(whitelist.WhiteName))
                        {
                            isSkip = whitelist.IsSkipConfirmation;
                        }
                    }
                }

                _checkList.BccAddresses.Add(new Address { MailAddress = i.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddresslist.Count == 0 || !alertAddresslist.Any(address => i.Key.Contains(address.TartgetAddress))) continue;

                _checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.IsAlertAddressBccAlert + $"[{i.Value}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                //送信禁止アドレスに該当する場合、禁止フラグを立て対象メールアドレスを説明文へ追加。
                foreach (var alertAddress in alertAddresslist)
                {
                    if (!i.Key.Contains(alertAddress.TartgetAddress) || !alertAddress.IsCanNotSend) continue;

                    _checkList.IsCanNotSendMail = true;
                    _checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{i.Value}]";
                }
            }
        }

        /// <summary>
        /// 送信遅延時間を算出する。
        /// 条件に該当する最も長い送信遅延時間を返す。
        /// </summary>
        /// <returns>送信遅延時間(分)</returns>
        private int CalcDeferredMinutes()
        {
            var readCsv = new ReadAndWriteCsv("DeferredDeliveryMinutes.csv");
            var deferredDeliveryMinutes = readCsv.GetCsvRecords<DeferredDeliveryMinutes>(readCsv.LoadCsv<DeferredDeliveryMinutesMap>());
            if (deferredDeliveryMinutes.Count == 0) return 0;

            var deferredMinutes = 0;

            //@のみで登録していた場合、それを標準の送信遅延時間とする。
            foreach (var config in deferredDeliveryMinutes)
            {
                if (config.TartgetAddress == "@")
                {
                    deferredMinutes = config.DeferredMinutes;
                }
            }

            if (_toDisplayNameAndRecipient.Count != 0)
            {
                foreach (var toRecipients in _toDisplayNameAndRecipient)
                {
                    foreach (var config in deferredDeliveryMinutes)
                    {
                        if (toRecipients.Value.Contains(config.TartgetAddress) && deferredMinutes < config.DeferredMinutes)
                        {
                            deferredMinutes = config.DeferredMinutes;
                        }
                    }
                }
            }

            if (_ccDisplayNameAndRecipient.Count != 0)
            {
                foreach (var ccRecipients in _ccDisplayNameAndRecipient)
                {
                    foreach (var config in deferredDeliveryMinutes)
                    {
                        if (ccRecipients.Value.Contains(config.TartgetAddress) && deferredMinutes < config.DeferredMinutes)
                        {
                            deferredMinutes = config.DeferredMinutes;
                        }
                    }
                }
            }

            if (_bccDisplayNameAndRecipient.Count != 0)
            {
                foreach (var bccRecipients in _bccDisplayNameAndRecipient)
                {
                    foreach (var config in deferredDeliveryMinutes)
                    {
                        if (bccRecipients.Value.Contains(config.TartgetAddress) && deferredMinutes < config.DeferredMinutes)
                        {
                            deferredMinutes = config.DeferredMinutes;
                        }
                    }
                }
            }

            return deferredMinutes;
        }
    }
}