using OutlookOkan.CsvTools;
using OutlookOkan.Properties;
using OutlookOkan.Types;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Models
{
    /// <summary>
    /// チェックリスト生成。
    /// </summary>
    public sealed class GenerateCheckList
    {
        private CheckList _checkList = new CheckList();
        private readonly List<Whitelist> _whitelist = new List<Whitelist>();

        /// <summary>
        /// メール送信の確認画面を表示。
        /// </summary>
        /// <param name="mail">送信するメールアイテム</param>
        /// <param name="generalSetting">一般設定</param>
        public CheckList GenerateCheckListFromMail(Outlook._MailItem mail, GeneralSetting generalSetting)
        {
            var whitelistCsv = new ReadAndWriteCsv("Whitelist.csv");
            _whitelist.AddRange(whitelistCsv.GetCsvRecords<Whitelist>(whitelistCsv.LoadCsv<WhitelistMap>()).Where(x => !string.IsNullOrEmpty(x.WhiteName)));

            var alertKeywordAndMessageListCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            var alertKeywordAndMessageList = alertKeywordAndMessageListCsv.GetCsvRecords<AlertKeywordAndMessage>(alertKeywordAndMessageListCsv.LoadCsv<AlertKeywordAndMessageMap>())
                .Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).ToList();

            var autoCcBccKeywordListCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");
            var autoCcBccKeywordList = autoCcBccKeywordListCsv.GetCsvRecords<AutoCcBccKeyword>(autoCcBccKeywordListCsv.LoadCsv<AutoCcBccKeywordMap>())
                .Where(x => !string.IsNullOrEmpty(x.AutoAddAddress) && !string.IsNullOrEmpty(x.Keyword)).ToList();

            var autoCcBccAttachedFilesListCsv = new ReadAndWriteCsv("AutoCcBccAttachedFileList.csv");
            var autoCcBccAttachedFilesList = autoCcBccAttachedFilesListCsv.GetCsvRecords<AutoCcBccAttachedFile>(autoCcBccAttachedFilesListCsv.LoadCsv<AutoCcBccAttachedFileMap>())
                .Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)).ToList();

            var autoCcBccRecipientListCsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");
            var autoCcBccRecipientList = autoCcBccRecipientListCsv.GetCsvRecords<AutoCcBccRecipient>(autoCcBccRecipientListCsv.LoadCsv<AutoCcBccRecipientMap>())
                .Where(x => !string.IsNullOrEmpty(x.AutoAddAddress) && !string.IsNullOrEmpty(x.TargetRecipient)).ToList();

            var alertAddressCsv = new ReadAndWriteCsv("AlertAddressList.csv");
            var alertAddressList = alertAddressCsv.GetCsvRecords<AlertAddress>(alertAddressCsv.LoadCsv<AlertAddressMap>())
                .Where(x => !string.IsNullOrEmpty(x.TargetAddress)).ToList();

            var nameAndDomainsCsv = new ReadAndWriteCsv("NameAndDomains.csv");
            var nameAndDomainsList = nameAndDomainsCsv.GetCsvRecords<NameAndDomains>(nameAndDomainsCsv.LoadCsv<NameAndDomainsMap>())
                .Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)).ToList();

            var deferredDeliveryMinutesCsv = new ReadAndWriteCsv("DeferredDeliveryMinutes.csv");
            var deferredDeliveryMinutes = deferredDeliveryMinutesCsv.GetCsvRecords<DeferredDeliveryMinutes>(deferredDeliveryMinutesCsv.LoadCsv<DeferredDeliveryMinutesMap>())
                .Where(x => !string.IsNullOrEmpty(x.TargetAddress)).ToList();

            var internalDomainListCsv = new ReadAndWriteCsv("InternalDomainList.csv");
            var internalDomainList = internalDomainListCsv.GetCsvRecords<InternalDomain>(internalDomainListCsv.LoadCsv<InternalDomainMap>())
                .Where(x => !string.IsNullOrEmpty(x.Domain)).ToList();

            var externalDomainsWarningAndAutoChangeToBccSetting = new ExternalDomainsWarningAndAutoChangeToBcc();
            var externalDomainsWarningAndAutoChangeToBccCsv = new ReadAndWriteCsv("ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
            var externalDomainsWarningAndAutoChangeToBccSettingList = externalDomainsWarningAndAutoChangeToBccCsv.GetCsvRecords<ExternalDomainsWarningAndAutoChangeToBcc>(externalDomainsWarningAndAutoChangeToBccCsv.LoadCsv<ExternalDomainsWarningAndAutoChangeToBccMap>());
            if (externalDomainsWarningAndAutoChangeToBccSettingList.Count > 0) externalDomainsWarningAndAutoChangeToBccSetting = externalDomainsWarningAndAutoChangeToBccSettingList[0];

            _checkList.Subject = mail.Subject ?? Resources.FailedToGetInformation;
            _checkList.MailType = GetMailBodyFormat(mail.BodyFormat) ?? Resources.FailedToGetInformation;
            _checkList.MailBody = GetMailBody(mail.BodyFormat, mail.Body ?? Resources.FailedToGetInformation);
            _checkList.MailHtmlBody = mail.HTMLBody ?? Resources.FailedToGetInformation;

            _checkList = GetSenderAndSenderDomain(in mail, _checkList);

            _checkList = GetAttachmentsInformation(in mail, _checkList, generalSetting.IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles);

            var displayNameAndRecipient = MakeDisplayNameAndRecipient(mail.Recipients, new DisplayNameAndRecipient(), generalSetting);

            _checkList = CheckForgotAttach(_checkList, generalSetting);

            _checkList = CheckKeyword(_checkList, alertKeywordAndMessageList);

            var autoAddRecipients = AutoAddCcAndBcc(mail, generalSetting, displayNameAndRecipient, autoCcBccKeywordList, autoCcBccAttachedFilesList, autoCcBccRecipientList, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, false));
            if (autoAddRecipients?.Count > 0)
            {
                displayNameAndRecipient = MakeDisplayNameAndRecipient(autoAddRecipients, displayNameAndRecipient, generalSetting);
                mail.Recipients.ResolveAll();
            }

            displayNameAndRecipient = ExternalDomainsChangeToBccIfNeeded(mail, displayNameAndRecipient,
                externalDomainsWarningAndAutoChangeToBccSetting, internalDomainList,
                CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, true),
                _checkList.SenderDomain, _checkList.Sender);

            _checkList = GetRecipient(_checkList, displayNameAndRecipient, alertAddressList, internalDomainList);

            _checkList = CheckMailBodyAndRecipient(_checkList, displayNameAndRecipient, nameAndDomainsList);

            _checkList.RecipientExternalDomainNumAll = CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, false);

            _checkList = ExternalDomainsWarningIfNeeded(_checkList, externalDomainsWarningAndAutoChangeToBccSetting, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, true));

            _checkList.DeferredMinutes = CalcDeferredMinutes(displayNameAndRecipient, deferredDeliveryMinutes, generalSetting.IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain, _checkList.RecipientExternalDomainNumAll);

            return _checkList;
        }

        /// <summary>
        /// 送信元アドレスと送信元ドメインを取得。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <param name="checkList">CheckList</param>
        /// <returns>CheckList</returns>
        private CheckList GetSenderAndSenderDomain(in Outlook._MailItem mail, CheckList checkList)
        {
            if (string.IsNullOrEmpty(mail.SentOnBehalfOfName))
            {
                checkList.Sender = mail.SendUsingAccount?.SmtpAddress ?? Resources.FailedToGetInformation;

                if (mail.SenderEmailType == "EX" && !IsValidEmailAddress(checkList.Sender))
                {
                    var tempOutlookApp = new Outlook.Application();
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(mail.SenderEmailAddress);

                    try
                    {
                        tempRecipient.Resolve();
                        var addressEntry = tempRecipient.AddressEntry;

                        var isDone = false;
                        var errorCount = 0;
                        while (!isDone && errorCount < 100)
                        {
                            try
                            {
                                var exchangeUser = addressEntry?.GetExchangeUser();
                                if (!(exchangeUser is null)) checkList.Sender = exchangeUser.PrimarySmtpAddress ?? Resources.FailedToGetInformation;

                                isDone = true;
                            }
                            catch (COMException e)
                            {
                                if (e.ErrorCode == -2147467260)
                                {
                                    //HRESULT:0x80004004 対策
                                    Thread.Sleep(10);
                                    errorCount++;
                                }
                                else
                                {
                                    isDone = true;
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }
                }
                else
                {
                    if (!IsValidEmailAddress(checkList.Sender))
                    {
                        checkList.Sender = mail.SenderEmailAddress ?? Resources.FailedToGetInformation;
                    }
                }

                if (!IsValidEmailAddress(checkList.Sender))
                {
                    checkList.Sender = Resources.FailedToGetInformation;
                }

                checkList.SenderDomain = checkList.Sender == Resources.FailedToGetInformation ? "------------------" : checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
            }
            else
            {
                //代理送信の場合。
                checkList.Sender = mail.Sender?.Address ?? Resources.FailedToGetInformation;

                if (IsValidEmailAddress(checkList.Sender))
                {
                    //メールアドレスが取得できる場合はそのまま使う。
                    checkList.SenderDomain = checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                    checkList.Sender = $@"{checkList.Sender} ([{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                }
                else
                {
                    //代理送信の場合かつExchangeのCN。
                    checkList.Sender = $@"[{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf}";
                    checkList.SenderDomain = @"------------------";

                    Outlook.ExchangeDistributionList exchangeDistributionList = null;
                    Outlook.ExchangeUser exchangeUser = null;

                    var sender = mail.Sender;

                    var isDone = false;
                    var errorCount = 0;
                    while (!isDone && errorCount < 100)
                    {
                        try
                        {
                            exchangeDistributionList = sender?.GetExchangeDistributionList();
                            exchangeUser = sender?.GetExchangeUser();

                            isDone = true;
                        }
                        catch (COMException e)
                        {
                            if (e.ErrorCode == -2147467260)
                            {
                                //HRESULT:0x80004004 対策
                                Thread.Sleep(10);
                                errorCount++;
                            }
                            else
                            {
                                isDone = true;
                            }
                        }
                    }

                    if (!(exchangeUser is null))
                    {
                        //ユーザの代理送信。
                        checkList.Sender = $@"{exchangeUser.PrimarySmtpAddress} ([{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                        checkList.SenderDomain = exchangeUser.PrimarySmtpAddress.Substring(exchangeUser.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                    }

                    if (!(exchangeDistributionList is null))
                    {
                        //配布リストの代理送信。
                        checkList.Sender = $@"{exchangeDistributionList.PrimarySmtpAddress} ([{mail.SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                        checkList.SenderDomain = exchangeDistributionList.PrimarySmtpAddress.Substring(exchangeDistributionList.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                    }
                }
            }

            return checkList;
        }

        /// <summary>
        /// メール本文をテキスト形式で取得する。
        /// </summary>
        /// <param name="mailBodyFormat">メール本文の種別</param>
        /// <param name="mailBody">メール本文</param>
        /// <returns>メール本文(テキスト形式)</returns>
        private string GetMailBody(Outlook.OlBodyFormat mailBodyFormat, string mailBody)
        {
            //改行が2行になる問題を回避するため、HTML形式の場合にのみ2行連続した改行を1行に置換する。
            return mailBodyFormat == Outlook.OlBodyFormat.olFormatHTML ? mailBody.Replace("\r\n\r\n", "\r\n") : mailBody;
        }

        /// <summary>
        /// 内部ドメインを除く宛先のドメイン数を数える。
        /// </summary>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="senderDomain">送信元ドメイン</param>
        /// <param name="internalDomain">内部ドメイン設定</param>
        /// <param name="isToAndCcOnly">対象がToとCcのみか否か</param>
        /// <returns>内部ドメインを除く宛先のドメイン数</returns>
        private int CountRecipientExternalDomains(DisplayNameAndRecipient displayNameAndRecipient, string senderDomain, IEnumerable<InternalDomain> internalDomain, bool isToAndCcOnly)
        {
            var domainList = new HashSet<string>();

            if (isToAndCcOnly)
            {
                foreach (var recipient in displayNameAndRecipient.To.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }

                foreach (var recipient in displayNameAndRecipient.Cc.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }
            else
            {
                foreach (var recipient in displayNameAndRecipient.All.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }

            var externalDomainsCount = domainList.Count;

            foreach (var _ in internalDomain.Where(settings => domainList.Contains(settings.Domain) && settings.Domain != senderDomain))
            {
                externalDomainsCount--;
            }

            //外部ドメインの数のため、送信者のドメインが含まれていた場合それをマイナスする。
            if (domainList.Contains(senderDomain))
            {
                return externalDomainsCount - 1;
            }

            return externalDomainsCount;
        }

        /// <summary>
        /// 宛先メールアドレスと宛先名称を取得する。
        /// </summary>
        /// <param name="recipient">メールの宛先</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetNameAndRecipient(Outlook.Recipient recipient)
        {
            var mailAddress = Resources.FailedToGetInformation;
            if (IsValidEmailAddress(recipient.Name))
            {
                mailAddress = recipient.Name;
            }
            else
            {
                if (IsValidEmailAddress(recipient.Address)) mailAddress = recipient.Address;
            }

            if (!IsValidEmailAddress(mailAddress))
            {
                try
                {
                    var propertyAccessor = recipient.PropertyAccessor;
                    Thread.Sleep(20);

                    var isDone = false;
                    var errorCount = 0;
                    while (!isDone && errorCount < 100)
                    {
                        try
                        {
                            mailAddress = propertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation;

                            isDone = true;
                        }
                        catch (COMException e)
                        {
                            if (e.ErrorCode == -2147467260)
                            {
                                //HRESULT:0x80004004 対策
                                Thread.Sleep(10);
                                errorCount++;
                            }
                            else
                            {
                                isDone = true;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // Do Nothing.
                }
            }

            if (!IsValidEmailAddress(mailAddress))
            {
                var tempOutlookApp = new Outlook.Application();
                var tempRecipient = tempOutlookApp.Session.CreateRecipient(recipient.Address);

                try
                {
                    recipient.Resolve();
                    var propertyAccessor = tempRecipient.AddressEntry.PropertyAccessor;
                    Thread.Sleep(20);

                    var isDone = false;
                    var errorCount = 0;
                    while (!isDone && errorCount < 100)
                    {
                        try
                        {
                            mailAddress = propertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation;
                            isDone = true;
                        }
                        catch (COMException e)
                        {
                            if (e.ErrorCode == -2147467260)
                            {
                                //HRESULT:0x80004004 対策
                                Thread.Sleep(10);
                                errorCount++;
                            }
                            else
                            {
                                isDone = true;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //Do Nothing.
                }
            }

            string nameAndMailAddress;
            if (string.IsNullOrEmpty(recipient.Name))
            {
                nameAndMailAddress = mailAddress ?? Resources.FailedToGetInformation;
            }
            else
            {
                nameAndMailAddress = recipient.Name.Contains($@" ({mailAddress})") ? recipient.Name : recipient.Name + $@" ({mailAddress})";
            }

            if (!IsValidEmailAddress(mailAddress)) mailAddress = nameAndMailAddress;

            return new List<NameAndRecipient> { new NameAndRecipient { MailAddress = mailAddress, NameAndMailAddress = nameAndMailAddress } };
        }

        /// <summary>
        /// Exchangeの配布リストを展開して宛先メールアドレスと宛先名称を取得する。(入れ子は非展開)
        /// </summary>
        /// <param name="recipient">メールの宛先</param>
        /// <param name="enableGetExchangeDistributionListMembers">配布リスト展開のオンオフ設定</param>
        /// <param name="exchangeDistributionListMembersAreWhite">配布リストで展開したアドレスをホワイトリスト化するか否かの設定</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetExchangeDistributionListMembers(Outlook.Recipient recipient, bool enableGetExchangeDistributionListMembers, bool exchangeDistributionListMembersAreWhite)
        {
            Outlook.OlAddressEntryUserType recipientAddressEntryUserType;
            try
            {
                recipientAddressEntryUserType = recipient.AddressEntry.AddressEntryUserType;
            }
            catch (Exception)
            {
                return null;
            }

            if (recipientAddressEntryUserType != Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry) return null;

            Outlook.ExchangeDistributionList distributionList = null;
            Outlook.AddressEntries addressEntries = null;

            try
            {
                var addressEntry = recipient.AddressEntry;

                var isDone = false;
                var errorCount = 0;
                while (!isDone && errorCount < 100)
                {
                    try
                    {
                        distributionList = addressEntry?.GetExchangeDistributionList();

                        if (enableGetExchangeDistributionListMembers)
                        {
                            addressEntries = distributionList?.GetExchangeDistributionListMembers();
                        }

                        isDone = true;
                    }
                    catch (COMException e)
                    {
                        if (e.ErrorCode == -2147467260)
                        {
                            //HRESULT:0x80004004 対策
                            Thread.Sleep(10);
                            errorCount++;
                        }
                        else
                        {
                            isDone = true;
                        }
                    }
                }

                if (distributionList is null) return null;

                var exchangeDistributionListMembers = new List<NameAndRecipient>();

                if (addressEntries is null || addressEntries.Count == 0)
                {
                    exchangeDistributionListMembers.Add(new NameAndRecipient { MailAddress = distributionList.PrimarySmtpAddress ?? Resources.FailedToGetInformation, NameAndMailAddress = (distributionList.Name ?? Resources.FailedToGetInformation) + $@" ({distributionList.PrimarySmtpAddress ?? Resources.FailedToGetInformation})" });

                    return exchangeDistributionListMembers;
                }

                var tempOutlookApp = new Outlook.Application();
                foreach (Outlook.AddressEntry member in addressEntries)
                {
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(member.Address);
                    var mailAddress = Resources.FailedToGetInformation;

                    try
                    {
                        tempRecipient.Resolve();
                        var propertyAccessor = tempRecipient.AddressEntry.PropertyAccessor;
                        Thread.Sleep(20);

                        isDone = false;
                        errorCount = 0;
                        while (!isDone && errorCount < 100)
                        {
                            try
                            {
                                mailAddress = propertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation;
                                isDone = true;
                            }
                            catch (COMException e)
                            {
                                if (e.ErrorCode == -2147467260)
                                {
                                    //HRESULT:0x80004004 対策
                                    Thread.Sleep(10);
                                    errorCount++;
                                }
                                else
                                {
                                    isDone = true;
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }

                    // 入れ子になった配布リストは、Exchangeサーバへの負荷が大きく、取得に時間もかかるため展開しない。
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
        /// 連絡先グループを展開して宛先メールアドレスと宛先名称を取得する。(入れ子も自動展開。)
        /// </summary>
        /// <param name="recipient">メールの宛先</param>
        /// <param name="contactGroupId">既に確認したGroupID</param>
        /// <param name="enableGetContactGroupMembers">連絡先グループ展開のオンオフ設定</param>
        /// <param name="contactGroupMembersAreWhite">連絡先グループで展開したアドレスをホワイトリスト化するか否かの設定</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetContactGroupMembers(Outlook.Recipient recipient, string contactGroupId, bool enableGetContactGroupMembers, bool contactGroupMembersAreWhite)
        {
            var contactGroupMembers = new List<NameAndRecipient>();
            if (!enableGetContactGroupMembers)
            {
                contactGroupMembers.Add(new NameAndRecipient { MailAddress = recipient.Name, NameAndMailAddress = recipient.Name + $@" [{Resources.ContactGroup}]" });
                return contactGroupMembers;
            }

            string entryId;
            if (contactGroupId is null)
            {
                var entryIdLength = Convert.ToInt32(recipient.AddressEntry.ID.Substring(66, 2) + recipient.AddressEntry.ID.Substring(64, 2), 16) * 2;
                entryId = recipient.AddressEntry.ID.Substring(72, entryIdLength);
            }
            else
            {
                //入れ子の場合のID。
                entryId = recipient.AddressEntry.ID.Substring(42);
            }

            if (contactGroupId?.Contains(entryId) == true) return null;

            contactGroupId = contactGroupId + entryId + ",";

            var tempOutlookApp = new Outlook.Application().GetNamespace("MAPI");
            var distList = (Outlook.DistListItem)tempOutlookApp.GetItemFromID(entryId);

            for (var i = 1; i < distList.MemberCount + 1; i++)
            {
                var member = distList.GetMember(i);
                contactGroupMembers.AddRange(member.Address == "Unknown"
                    ? GetContactGroupMembers(member, contactGroupId, true, contactGroupMembersAreWhite)
                    : GetNameAndRecipient(member));
            }

            foreach (var nameAndRecipient in contactGroupMembers)
            {
                nameAndRecipient.IncludedGroupAndList += $@" [{distList.DLName}]";

                if (contactGroupMembersAreWhite)
                {
                    _whitelist.Add(new Whitelist { WhiteName = nameAndRecipient.MailAddress });
                }
            }

            return contactGroupMembers;
        }

        /// <summary>
        /// 送信先の表示名と表示名とメールアドレスを対応させる。(Outlookの仕様上、表示名にメールアドレスが含まれない事がある。)
        /// </summary>
        /// <param name="recipients">メールの宛先</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="generalSetting">一般設定</param>
        /// <returns>宛先アドレスと名称</returns>
        private DisplayNameAndRecipient MakeDisplayNameAndRecipient(IEnumerable recipients, DisplayNameAndRecipient displayNameAndRecipient, GeneralSetting generalSetting)
        {
            foreach (Outlook.Recipient recipient in recipients)
            {
                var recipientAddressEntryUserType = Outlook.OlAddressEntryUserType.olOtherAddressEntry;
                try
                {
                    recipientAddressEntryUserType = recipient.AddressEntry.AddressEntryUserType;
                }
                catch (Exception)
                {
                    //Do Nothing.
                }

                var nameAndRecipients = new List<NameAndRecipient>();

                switch (recipientAddressEntryUserType)
                {
                    case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                        var exchangeMembers = GetExchangeDistributionListMembers(recipient, generalSetting.EnableGetExchangeDistributionListMembers, generalSetting.ExchangeDistributionListMembersAreWhite);
                        if (exchangeMembers is null)
                        {
                            nameAndRecipients.AddRange(GetNameAndRecipient(recipient));
                            break;
                        }
                        else
                        {
                            nameAndRecipients.AddRange(exchangeMembers);
                            break;
                        }
                    case Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry:
                        var addressEntryMembers = GetContactGroupMembers(recipient, null, generalSetting.EnableGetContactGroupMembers, generalSetting.ContactGroupMembersAreWhite);
                        if (addressEntryMembers is null)
                        {
                            nameAndRecipients.AddRange(GetNameAndRecipient(recipient));
                            break;
                        }
                        else
                        {
                            nameAndRecipients.AddRange(addressEntryMembers);
                            break;
                        }
                    default:
                        nameAndRecipients.AddRange(GetNameAndRecipient(recipient));
                        break;
                }

                foreach (var nameAndRecipient in nameAndRecipients)
                {
                    if (displayNameAndRecipient.All.ContainsKey(nameAndRecipient.MailAddress))
                    {
                        displayNameAndRecipient.All[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                    }
                    else
                    {
                        displayNameAndRecipient.All[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                    }

                    displayNameAndRecipient.MailRecipientsIndex.Add(new MailItemsRecipientAndMailAddress
                    {
                        MailAddress = nameAndRecipient.MailAddress,
                        MailItemsRecipient = recipient.Address,
                        Type = recipient.Type
                    });

                    switch (recipient.Type)
                    {
                        case 1:
                            if (displayNameAndRecipient.To.ContainsKey(nameAndRecipient.MailAddress))
                            {
                                displayNameAndRecipient.To[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                            }
                            else
                            {
                                displayNameAndRecipient.To[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                            }
                            continue;
                        case 2:
                            if (displayNameAndRecipient.Cc.ContainsKey(nameAndRecipient.MailAddress))
                            {
                                displayNameAndRecipient.Cc[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                            }
                            else
                            {
                                displayNameAndRecipient.Cc[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                            }
                            continue;
                        case 3:
                            if (displayNameAndRecipient.Bcc.ContainsKey(nameAndRecipient.MailAddress))
                            {
                                displayNameAndRecipient.Bcc[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                            }
                            else
                            {
                                displayNameAndRecipient.Bcc[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                            }
                            continue;
                        default:
                            continue;
                    }
                }
            }

            return displayNameAndRecipient;
        }

        /// <summary>
        /// ファイルの添付忘れを確認。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="generalSetting">一般設定</param>
        /// <returns>CheckList</returns>
        private CheckList CheckForgotAttach(CheckList checkList, GeneralSetting generalSetting)
        {
            if (checkList.Attachments.Count >= 1) return checkList;

            if (!generalSetting.EnableForgottenToAttachAlert) return checkList;

            string attachmentsKeyword;
            switch (generalSetting.LanguageCode)
            {
                case "ja-JP":
                    attachmentsKeyword = "添付";
                    break;
                case "en-US":
                    attachmentsKeyword = "attached file";
                    break;
                default:
                    //設定値がなければ、日本語環境として扱う。
                    attachmentsKeyword = "添付";
                    break;
            }

            if (checkList.MailBody.Contains(attachmentsKeyword))
            {
                checkList.Alerts.Add(new Alert { AlertMessage = Resources.ForgottenToAttachAlert, IsImportant = true, IsWhite = false, IsChecked = false });
            }

            return checkList;
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
        /// <param name="checkList">CheckList</param>>
        /// <param name="alertKeywordAndMessageList">警告キーワード設定</param>>
        /// <returns>CheckList</returns>
        private CheckList CheckKeyword(CheckList checkList, IReadOnlyCollection<AlertKeywordAndMessage> alertKeywordAndMessageList)
        {
            if (alertKeywordAndMessageList.Count == 0) return checkList;

            foreach (var alertKeywordAndMessage in alertKeywordAndMessageList)
            {
                if (!checkList.MailBody.Contains(alertKeywordAndMessage.AlertKeyword)) continue;

                var alertMessage = Resources.DefaultAlertMessage + $"[{alertKeywordAndMessage.AlertKeyword}]";
                if (!string.IsNullOrEmpty(alertKeywordAndMessage.Message)) alertMessage = alertKeywordAndMessage.Message;

                checkList.Alerts.Add(new Alert { AlertMessage = alertMessage, IsImportant = true, IsWhite = false, IsChecked = false });

                if (!alertKeywordAndMessage.IsCanNotSend) continue;

                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = alertMessage;
            }

            return checkList;
        }

        /// <summary>
        /// 条件に一致した場合、CcやBccに宛先を追加する。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <param name="generalSetting">一般設定</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称設定</param>
        /// <param name="autoCcBccKeywordList">自動Cc/Bcc追加(キーワード)設定</param>
        /// <param name="autoCcBccAttachedFilesList">自動Cc/Bcc追加(キーワード)設定</param>
        /// <param name="autoCcBccRecipientList">自動Cc/Bcc追加(宛先)設定</param>
        /// <param name="externalDomainCount">外部ドメイン数</param>
        /// <returns>CcやBccに自動追加された宛先アドレス</returns>
        private List<Outlook.Recipient> AutoAddCcAndBcc(Outlook._MailItem mail, GeneralSetting generalSetting, DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<AutoCcBccKeyword> autoCcBccKeywordList, IReadOnlyCollection<AutoCcBccAttachedFile> autoCcBccAttachedFilesList, IReadOnlyCollection<AutoCcBccRecipient> autoCcBccRecipientList, int externalDomainCount)
        {
            var autoAddedCcAddressList = new List<string>();
            var autoAddedBccAddressList = new List<string>();
            var autoAddRecipients = new List<Outlook.Recipient>();

            if (autoCcBccKeywordList.Count != 0 && !(externalDomainCount == 0 && generalSetting.IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain))
            {
                foreach (var autoCcBccKeyword in autoCcBccKeywordList)
                {
                    if (!_checkList.MailBody.Contains(autoCcBccKeyword.Keyword) || !autoCcBccKeyword.AutoAddAddress.Contains("@")) continue;

                    if (autoCcBccKeyword.CcOrBcc == CcOrBcc.Cc)
                    {
                        if (!autoAddedCcAddressList.Contains(autoCcBccKeyword.AutoAddAddress) && !displayNameAndRecipient.Cc.ContainsKey(autoCcBccKeyword.AutoAddAddress))
                        {
                            var recipient = mail.Recipients.Add(autoCcBccKeyword.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedCcAddressList.Add(autoCcBccKeyword.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(autoCcBccKeyword.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccKeyword.AutoAddAddress))
                    {
                        var recipient = mail.Recipients.Add(autoCcBccKeyword.AutoAddAddress);
                        recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recipient);
                        autoAddedBccAddressList.Add(autoCcBccKeyword.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccKeyword.CcOrBcc}] [{autoCcBccKeyword.AutoAddAddress}] (" + Resources.ApplicableKeywords + $" 「{autoCcBccKeyword.Keyword}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    _whitelist.Add(new Whitelist { WhiteName = autoCcBccKeyword.AutoAddAddress });
                }
            }

            //警告対象の添付ファイル数が0でない場合のみ、CcやBccの追加処理を行う。
            if (_checkList.Attachments.Count != 0 && !(externalDomainCount == 0 && generalSetting.IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain))
            {
                if (autoCcBccAttachedFilesList.Count != 0)
                {
                    foreach (var autoCcBccAttachedFile in autoCcBccAttachedFilesList)
                    {
                        if (autoCcBccAttachedFile.CcOrBcc == CcOrBcc.Cc)
                        {
                            if (!autoAddedCcAddressList.Contains(autoCcBccAttachedFile.AutoAddAddress) && !displayNameAndRecipient.Cc.ContainsKey(autoCcBccAttachedFile.AutoAddAddress))
                            {
                                var recipient = mail.Recipients.Add(autoCcBccAttachedFile.AutoAddAddress);
                                recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                                autoAddRecipients.Add(recipient);
                                autoAddedCcAddressList.Add(autoCcBccAttachedFile.AutoAddAddress);
                            }
                        }
                        else if (!autoAddedBccAddressList.Contains(autoCcBccAttachedFile.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccAttachedFile.AutoAddAddress))
                        {
                            var recipient = mail.Recipients.Add(autoCcBccAttachedFile.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedBccAddressList.Add(autoCcBccAttachedFile.AutoAddAddress);
                        }

                        _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccAttachedFile.CcOrBcc}] [{autoCcBccAttachedFile.AutoAddAddress}] (" + Resources.Attachments + ")", IsImportant = false, IsWhite = true, IsChecked = true });

                        _whitelist.Add(new Whitelist { WhiteName = autoCcBccAttachedFile.AutoAddAddress });
                    }
                }
            }

            if (autoCcBccRecipientList.Count != 0)
            {
                foreach (var autoCcBccRecipient in autoCcBccRecipientList)
                {
                    if (!displayNameAndRecipient.All.Any(recipient => recipient.Key.Contains(autoCcBccRecipient.TargetRecipient)) || !autoCcBccRecipient.AutoAddAddress.Contains("@")) continue;

                    if (autoCcBccRecipient.CcOrBcc == CcOrBcc.Cc)
                    {
                        if (!autoAddedCcAddressList.Contains(autoCcBccRecipient.AutoAddAddress) && !displayNameAndRecipient.Cc.ContainsKey(autoCcBccRecipient.AutoAddAddress))
                        {
                            var recipient = mail.Recipients.Add(autoCcBccRecipient.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedCcAddressList.Add(autoCcBccRecipient.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(autoCcBccRecipient.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccRecipient.AutoAddAddress))
                    {
                        var recipient = mail.Recipients.Add(autoCcBccRecipient.AutoAddAddress);
                        recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recipient);
                        autoAddedBccAddressList.Add(autoCcBccRecipient.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccRecipient.CcOrBcc}] [{autoCcBccRecipient.AutoAddAddress}] (" + Resources.ApplicableDestination + $" 「{autoCcBccRecipient.TargetRecipient}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    _whitelist.Add(new Whitelist { WhiteName = autoCcBccRecipient.AutoAddAddress });
                }
            }

            return autoAddRecipients;
        }

        /// <summary>
        /// HTML内に埋め込まれた添付ファイル名を取得する。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <returns>埋め込みファイル名のList</returns>
        private List<string> MakeEmbeddedAttachmentsList(in Outlook._MailItem mail)
        {
            //HTML形式の場合のみ、処理対象とする。
            if (mail.BodyFormat != Outlook.OlBodyFormat.olFormatHTML) return null;

            var htmlBody = mail.HTMLBody;
            var matches = Regex.Matches(htmlBody, @"cid:.*?@");

            if (matches.Count == 0) return null;

            var embeddedAttachmentsName = new List<string>();
            foreach (var data in matches)
            {
                embeddedAttachmentsName.Add(data.ToString().Replace(@"cid:", "").Replace(@"@", ""));
            }

            return embeddedAttachmentsName;
        }

        /// <summary>
        /// 添付ファイルとそのファイルサイズを取得し、チェックリストに追加する。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <param name="checkList">CheckList</param>
        /// <param name="isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles">HTML埋め込みの添付ファイル無視設定</param>
        /// <returns>CheckList</returns>
        private CheckList GetAttachmentsInformation(in Outlook._MailItem mail, CheckList checkList, bool isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles)
        {
            if (mail.Attachments.Count == 0) return checkList;

            var embeddedAttachmentsName = new List<string>();
            if (isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles)
            {
                embeddedAttachmentsName = MakeEmbeddedAttachmentsList(in mail);
            }

            for (var i = 0; i < mail.Attachments.Count; i++)
            {
                var fileSize = "?KB";
                if (mail.Attachments[i + 1].Size != 0)
                {
                    fileSize = Math.Round(((double)mail.Attachments[i + 1].Size / 1024), 0, MidpointRounding.AwayFromZero).ToString("##,###") + "KB";
                }

                if (mail.Attachments[i + 1].Size >= 10485760)
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsBigAttachedFile + $"[{mail.Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
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
                if (fileType == ".exe")
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsAttachedExe + $"[{mail.Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                    isDangerous = true;
                }

                string fileName;
                try
                {
                    fileName = mail.Attachments[i + 1].FileName;
                }
                catch (Exception)
                {
                    fileName = Resources.Unknown;
                }

                //情報取得に完全に失敗した添付ファイルは無視する。(リッチテキスト形式の埋め込み画像など)
                if (fileName == Resources.Unknown && fileSize == "?KB" && fileType == Resources.Unknown) continue;

                if (embeddedAttachmentsName is null)
                {
                    checkList.Attachments.Add(new Attachment
                    {
                        FileName = fileName,
                        FileSize = fileSize,
                        FileType = fileType,
                        IsTooBig = mail.Attachments[i + 1].Size >= 10485760,
                        IsEncrypted = false,
                        IsChecked = false,
                        IsDangerous = isDangerous
                    });

                    continue;
                }

                //HTML埋め込みファイルは無視する。
                if (!embeddedAttachmentsName.Contains(fileName))
                {
                    checkList.Attachments.Add(new Attachment
                    {
                        FileName = fileName,
                        FileSize = fileSize,
                        FileType = fileType,
                        IsTooBig = mail.Attachments[i + 1].Size >= 10485760,
                        IsEncrypted = false,
                        IsChecked = false,
                        IsDangerous = isDangerous
                    });
                }
            }

            return checkList;
        }

        /// <summary>
        /// 登録された名称とドメインから、宛先候補ではないアドレスが宛先に含まれている場合に、警告を表示する。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="nameAndDomainsList">宛先と名称のリスト</param>
        /// <returns>CheckList</returns>
        private CheckList CheckMailBodyAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<NameAndDomains> nameAndDomainsList)
        {
            if (displayNameAndRecipient is null) return checkList;

            //空の設定値があると誤検知するため、空を省く。
            var cleanedNameAndDomains = nameAndDomainsList.Where(nameAndDomain => !string.IsNullOrEmpty(nameAndDomain.Domain) && !string.IsNullOrEmpty(nameAndDomain.Name)).ToList();

            var recipientCandidateDomains = (from nameAndDomain in cleanedNameAndDomains where checkList.MailBody.Contains(nameAndDomain.Name) select nameAndDomain.Domain).ToList();

            //送信先の候補が見つからない場合、何もしない。(見つからない場合の方が多いため、警告ばかりになってしまう。)
            if (recipientCandidateDomains.Count == 0) return checkList;

            foreach (var recipient in displayNameAndRecipient.All)
            {
                if (recipientCandidateDomains.Any(domains => domains.Equals(recipient.Key.Substring(recipient.Key.IndexOf("@", StringComparison.Ordinal))))) continue;

                //送信者ドメインは警告対象外。
                if (!recipient.Key.Contains(checkList.SenderDomain))
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = recipient.Value + " : " + Resources.IsAlertAddressMaybeIrrelevant, IsImportant = true, IsWhite = false, IsChecked = false });
                }
            }

            return checkList;
        }

        /// <summary>
        /// 送信先メールアドレスを取得し、チェックリストに追加する。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="alertAddressList">警告アドレス設定</param>
        /// <param name="internalDomainList">内部ドメイン設定</param>
        /// <returns>CheckList</returns>
        private CheckList GetRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<AlertAddress> alertAddressList, List<InternalDomain> internalDomainList)
        {
            if (displayNameAndRecipient is null) return checkList;

            var internalDomains = internalDomainList;
            internalDomains.Add(new InternalDomain { Domain = checkList.SenderDomain });

            foreach (var to in displayNameAndRecipient.To)
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(settings => to.Key.Contains(settings.Domain)))
                {
                    isExternal = false;
                }

                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => to.Key.Contains(x.WhiteName));
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist.Where(whitelist => to.Key.Contains(whitelist.WhiteName)))
                    {
                        isSkip = whitelist.IsSkipConfirmation;
                    }
                }

                checkList.ToAddresses.Add(new Address { MailAddress = to.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddressList.Count == 0 || !alertAddressList.Any(address => to.Key.Contains(address.TargetAddress))) continue;

                checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.IsAlertAddressToAlert + $"[{to.Value}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                foreach (var alertAddress in alertAddressList)
                {
                    if (!to.Key.Contains(alertAddress.TargetAddress) || !alertAddress.IsCanNotSend) continue;

                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{to.Value}]";
                }
            }

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(settings => cc.Key.Contains(settings.Domain)))
                {
                    isExternal = false;
                }

                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => cc.Key.Contains(x.WhiteName));
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist.Where(whitelist => cc.Key.Contains(whitelist.WhiteName)))
                    {
                        isSkip = whitelist.IsSkipConfirmation;
                    }
                }

                checkList.CcAddresses.Add(new Address { MailAddress = cc.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddressList.Count == 0 || !alertAddressList.Any(address => cc.Key.Contains(address.TargetAddress))) continue;

                checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.IsAlertAddressCcAlert + $"[{cc.Value}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                foreach (var alertAddress in alertAddressList)
                {
                    if (!cc.Key.Contains(alertAddress.TargetAddress) || !alertAddress.IsCanNotSend) continue;

                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{cc.Value}]";
                }
            }

            foreach (var bcc in displayNameAndRecipient.Bcc)
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(settings => bcc.Key.Contains(settings.Domain)))
                {
                    isExternal = false;
                }

                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => bcc.Key.Contains(x.WhiteName));
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist.Where(whitelist => bcc.Key.Contains(whitelist.WhiteName)))
                    {
                        isSkip = whitelist.IsSkipConfirmation;
                    }
                }

                checkList.BccAddresses.Add(new Address { MailAddress = bcc.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddressList.Count == 0 || !alertAddressList.Any(address => bcc.Key.Contains(address.TargetAddress))) continue;

                checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.IsAlertAddressBccAlert + $"[{bcc.Value}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                foreach (var alertAddress in alertAddressList)
                {
                    if (!bcc.Key.Contains(alertAddress.TargetAddress) || !alertAddress.IsCanNotSend) continue;

                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{bcc.Value}]";
                }
            }

            return checkList;
        }

        /// <summary>
        /// 送信遅延時間を算出する。(条件に該当する最も長い送信遅延時間を返す。)
        /// </summary>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="deferredDeliveryMinutes">遅延送信設定</param>
        /// <param name="isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain">外部ドメイン数が0の際の機能使用有無</param>
        /// <param name="externalDomainCount">外部ドメイン数</param>
        /// <returns>送信遅延時間(分)</returns>
        private int CalcDeferredMinutes(DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<DeferredDeliveryMinutes> deferredDeliveryMinutes, bool isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain, int externalDomainCount)
        {
            if (deferredDeliveryMinutes.Count == 0) return 0;
            if (externalDomainCount == 0 && isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain) return 0;

            var deferredMinutes = 0;

            //@のみで登録していた場合、それを標準の送信遅延時間とする。
            foreach (var config in deferredDeliveryMinutes.Where(config => config.TargetAddress == "@"))
            {
                deferredMinutes = config.DeferredMinutes;
            }

            if (displayNameAndRecipient.To.Count != 0)
            {
                foreach (var toRecipients in displayNameAndRecipient.To)
                {
                    foreach (var config in deferredDeliveryMinutes.Where(config => toRecipients.Key.Contains(config.TargetAddress) && deferredMinutes < config.DeferredMinutes))
                    {
                        deferredMinutes = config.DeferredMinutes;
                    }
                }
            }

            if (displayNameAndRecipient.Cc.Count != 0)
            {
                foreach (var ccRecipients in displayNameAndRecipient.Cc)
                {
                    foreach (var config in deferredDeliveryMinutes.Where(config => ccRecipients.Key.Contains(config.TargetAddress) && deferredMinutes < config.DeferredMinutes))
                    {
                        deferredMinutes = config.DeferredMinutes;
                    }
                }
            }

            if (displayNameAndRecipient.Bcc.Count != 0)
            {
                foreach (var bccRecipients in displayNameAndRecipient.Bcc)
                {
                    foreach (var config in deferredDeliveryMinutes.Where(config => bccRecipients.Key.Contains(config.TargetAddress) && deferredMinutes < config.DeferredMinutes))
                    {
                        deferredMinutes = config.DeferredMinutes;
                    }
                }
            }

            return deferredMinutes;
        }

        /// <summary>
        /// ToとCcの外部ドメイン数が規定値以上の場合に、警告を表示したり、メール送信を禁止したりする。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="settings">外部ドメイン数の警告と自動Bcc変換の設定</param>
        /// <param name="externalDomainNumToAndCc">ToとCcの外部ドメイン数</param>
        /// <returns>CheckList</returns>
        private CheckList ExternalDomainsWarningIfNeeded(CheckList checkList, ExternalDomainsWarningAndAutoChangeToBcc settings, int externalDomainNumToAndCc)
        {
            if (settings.TargetToAndCcExternalDomainsNum > externalDomainNumToAndCc) return checkList;

            if (settings.IsProhibitedWhenLargeNumberOfExternalDomains)
            {
                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = Resources.ProhibitedWhenLargeNumberOfExternalDomainsAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]";

                return checkList;
            }

            if (settings.IsWarningWhenLargeNumberOfExternalDomains && !settings.IsAutoChangeToBccWhenLargeNumberOfExternalDomains)
            {
                checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.LargeNumberOfExternalDomainAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                return checkList;
            }

            return checkList;
        }

        /// <summary>
        /// 指定した宛先をTo及びCcから削除し、Bccへ追加する。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <param name="mailItemsRecipientAndMailAddress">メールアドレスとMailItemのRecipient</param>
        /// <param name="senderMailAddress">送信元メールアドレス</param>
        /// <param name="isNeedsAddToSender">Toへ送信元アドレスを追加するか否か</param>
        private void ChangeToBcc(Outlook._MailItem mail, IReadOnlyCollection<MailItemsRecipientAndMailAddress> mailItemsRecipientAndMailAddress, string senderMailAddress, bool isNeedsAddToSender)
        {
            if (mail is null) return;

            var targetMailAddressAndRecipient = new Dictionary<string, string>();

            foreach (Outlook.Recipient recipient in mail.Recipients)
            {
                foreach (var target in mailItemsRecipientAndMailAddress)
                {
                    if (recipient.Address == target.MailItemsRecipient) targetMailAddressAndRecipient[target.MailAddress] = target.MailItemsRecipient;
                }
            }

            //Indexを使用してRemoveした場合、Indexがずれ、複数を正しく削除できないため、削除対象を探して削除する。
            var targetCount = targetMailAddressAndRecipient.Count;
            while (targetCount != 0)
            {
                foreach (var target in targetMailAddressAndRecipient)
                {
                    foreach (Outlook.Recipient recipient in mail.Recipients)
                    {
                        if (recipient.Address != target.Value) continue;
                        mail.Recipients.Remove(recipient.Index);
                        targetCount--;
                    }
                }
            }

            foreach (var addTarget in targetMailAddressAndRecipient.Select(mailAddress => mail.Recipients.Add(mailAddress.Key)))
            {
                addTarget.Type = (int)Outlook.OlMailRecipientType.olBCC;
            }

            if (isNeedsAddToSender)
            {
                var senderRecipient = mail.Recipients.Add(senderMailAddress);
                senderRecipient.Type = (int)Outlook.OlMailRecipientType.olTo;
            }

            mail.Recipients.ResolveAll();
        }

        /// <summary>
        /// 条件に該当する場合、ToとCcの外部アドレスをBccに変換する。
        /// </summary>
        /// <param name="mail">Mail</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="settings">外部ドメイン数の警告と自動Bcc変換の設定</param>
        /// <param name="internalDomains">内部ドメイン</param>
        /// <param name="externalDomainNumToAndCc">ToとCcの外部ドメイン数</param>
        /// <param name="senderDomain">送信元ドメイン</param>
        /// <param name="senderMailAddress">送信元アドレス</param>
        /// <returns>DisplayNameAndRecipient</returns>
        private DisplayNameAndRecipient ExternalDomainsChangeToBccIfNeeded(Outlook._MailItem mail, DisplayNameAndRecipient displayNameAndRecipient, ExternalDomainsWarningAndAutoChangeToBcc settings, ICollection<InternalDomain> internalDomains, int externalDomainNumToAndCc, string senderDomain, string senderMailAddress)
        {
            if (!settings.IsAutoChangeToBccWhenLargeNumberOfExternalDomains || settings.IsProhibitedWhenLargeNumberOfExternalDomains || settings.TargetToAndCcExternalDomainsNum > externalDomainNumToAndCc) return displayNameAndRecipient;

            internalDomains.Add(new InternalDomain { Domain = senderDomain });
            var removeTarget = new List<string>();

            foreach (var to in displayNameAndRecipient.To)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => to.Key.Contains(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[to.Key] = to.Value;
                removeTarget.Add(to.Key);
            }
            foreach (var target in removeTarget)
            {
                displayNameAndRecipient.To.Remove(target);
            }

            removeTarget.Clear();

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => cc.Key.Contains(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[cc.Key] = cc.Value;
                removeTarget.Add(cc.Key);
            }
            foreach (var target in removeTarget)
            {
                displayNameAndRecipient.Cc.Remove(target);
            }

            AddAlerts(Resources.ExternalDomainsChangeToBccAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]", true, false, false);

            //Toが存在しなくなる場合、送信者をToに追加する。
            var isNeedsAddToSender = false;
            if (displayNameAndRecipient.To.Count == 0)
            {
                displayNameAndRecipient.To[senderMailAddress] = senderMailAddress;
                AddAlerts(Resources.AutoAddSendersAddressToAlert, true, false, false);
                isNeedsAddToSender = true;
            }

            var targetMailRecipientsIndex = new List<MailItemsRecipientAndMailAddress>();
            //元からBccの宛先には何もしない。
            foreach (var recipient in displayNameAndRecipient.MailRecipientsIndex.Where(recipient => recipient.Type != (int)Outlook.OlMailRecipientType.olBCC))
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(internalDomain => recipient.MailAddress.Contains(internalDomain.Domain)))
                {
                    isExternal = false;
                }

                if (isExternal) targetMailRecipientsIndex.Add(recipient);
            }

            ChangeToBcc(mail, targetMailRecipientsIndex, senderMailAddress, isNeedsAddToSender);

            return displayNameAndRecipient;
        }

        /// <summary>
        /// 警告を追加する。
        /// </summary>
        /// <param name="alertMessage">警告文</param>
        /// <param name="isImportant">重要か否か</param>
        /// <param name="isWhite">一旦未使用</param>
        /// <param name="isChecked">既定のチェック有無</param>
        private void AddAlerts(string alertMessage, bool isImportant, bool isWhite, bool isChecked)
        {
            _checkList.Alerts.Add(new Alert
            {
                AlertMessage = alertMessage,
                IsImportant = isImportant,
                IsWhite = isWhite,
                IsChecked = isChecked
            });
        }

        /// <summary>
        /// メールアドレスか否か判定する。
        /// </summary>
        /// <param name="emailAddress">判定したい文字列</param>
        /// <returns>メールアドレスか否か</returns>
        private bool IsValidEmailAddress(string emailAddress)
        {
            if (string.IsNullOrWhiteSpace(emailAddress)) return false;

            try
            {
                emailAddress = Regex.Replace(emailAddress, @"(@)(.+)$", DomainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(500));
                string DomainMapper(Match match)
                {
                    var idnMapping = new IdnMapping();
                    var domainName = idnMapping.GetAscii(match.Groups[2].Value);
                    return match.Groups[1].Value + domainName;
                }
            }
            catch (Exception)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(emailAddress, @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" + @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(500));
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}