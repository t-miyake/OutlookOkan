using OutlookOkan.CsvTools;
using OutlookOkan.Properties;
using OutlookOkan.Types;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
        private int _failedToGetInformationOfRecipientsMailAddressCounter;

        /// <summary>
        /// メール送信前の確認画面で使用するチェックリストの生成。
        /// </summary>
        /// <param name="item">送信するアイテム</param>
        /// <param name="generalSetting">一般設定</param>
        /// <param name="contacts">連絡先(アドレス帳)</param>
        /// <param name="autoAddMessageSetting">autoAddMessageSetting</param>
        internal CheckList GenerateCheckListFromMail<T>(T item, GeneralSetting generalSetting, Outlook.MAPIFolder contacts, AutoAddMessage autoAddMessageSetting)
        {
            #region LoadSettings

            var whitelistCsv = new ReadAndWriteCsv("Whitelist.csv");
            _whitelist.AddRange(whitelistCsv.GetCsvRecords<Whitelist>(whitelistCsv.LoadCsv<WhitelistMap>()).Where(x => !string.IsNullOrEmpty(x.WhiteName)));

            var alertKeywordAndMessageListCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            var alertKeywordAndMessageList = alertKeywordAndMessageListCsv.GetCsvRecords<AlertKeywordAndMessage>(alertKeywordAndMessageListCsv.LoadCsv<AlertKeywordAndMessageMap>())
                .Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).ToList();

            var alertKeywordAndMessageForSubjectListCsv = new ReadAndWriteCsv("AlertKeywordAndMessageListForSubject.csv");
            var alertKeywordAndMessageForSubjectList = alertKeywordAndMessageForSubjectListCsv.GetCsvRecords<AlertKeywordAndMessageForSubject>(alertKeywordAndMessageForSubjectListCsv.LoadCsv<AlertKeywordAndMessageForSubjectMap>())
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

            var attachmentsSetting = new AttachmentsSetting();
            var attachmentsSettingCsv = new ReadAndWriteCsv("AttachmentsSetting.csv");
            var attachmentsSettingList = attachmentsSettingCsv.GetCsvRecords<AttachmentsSetting>(attachmentsSettingCsv.LoadCsv<AttachmentsSettingMap>());
            if (attachmentsSettingList.Count > 0) attachmentsSetting = attachmentsSettingList[0];

            var recipientsAndAttachmentsNameCsv = new ReadAndWriteCsv("RecipientsAndAttachmentsName.csv");
            var recipientsAndAttachmentsNameList = recipientsAndAttachmentsNameCsv.GetCsvRecords<RecipientsAndAttachmentsName>(recipientsAndAttachmentsNameCsv.LoadCsv<RecipientsAndAttachmentsNameMap>())
                .Where(x => !string.IsNullOrEmpty(x.Recipient) && !string.IsNullOrEmpty(x.AttachmentsName)).ToList();

            var attachmentProhibitedRecipientsCsv = new ReadAndWriteCsv("AttachmentProhibitedRecipients.csv");
            var attachmentProhibitedRecipientsList = attachmentProhibitedRecipientsCsv.GetCsvRecords<AttachmentProhibitedRecipients>(attachmentProhibitedRecipientsCsv.LoadCsv<AttachmentProhibitedRecipientsMap>())
                .Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList();

            var attachmentAlertRecipientsCsv = new ReadAndWriteCsv("AttachmentAlertRecipients.csv");
            var attachmentAlertRecipientsList = attachmentAlertRecipientsCsv.GetCsvRecords<AttachmentAlertRecipients>(attachmentAlertRecipientsCsv.LoadCsv<AttachmentAlertRecipientsMap>())
                .Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList();

            var forceAutoChangeRecipientsToBccSetting = new ForceAutoChangeRecipientsToBcc();
            var forceAutoChangeRecipientsToBccCsv = new ReadAndWriteCsv("ForceAutoChangeRecipientsToBcc.csv");
            var forceAutoChangeRecipientsToBccSettingList = forceAutoChangeRecipientsToBccCsv.GetCsvRecords<ForceAutoChangeRecipientsToBcc>(forceAutoChangeRecipientsToBccCsv.LoadCsv<ForceAutoChangeRecipientsToBccMap>());
            if (forceAutoChangeRecipientsToBccSettingList.Count > 0) forceAutoChangeRecipientsToBccSetting = forceAutoChangeRecipientsToBccSettingList[0];

            #endregion

            var isMailItem = (typeof(T) == typeof(Outlook._MailItem));

            if (isMailItem)
            {
                _checkList.MailType = GetMailBodyFormat(((Outlook._MailItem)item).BodyFormat) ?? Resources.FailedToGetInformation;
                _checkList.MailBody = GetMailBody(((Outlook._MailItem)item).BodyFormat, ((Outlook._MailItem)item).Body ?? Resources.FailedToGetInformation);
                _checkList.MailBody = AddMessageToBodyPreview(_checkList.MailBody, autoAddMessageSetting);

                _checkList.MailHtmlBody = ((Outlook._MailItem)item).HTMLBody ?? Resources.FailedToGetInformation;
            }
            else
            {
                _checkList.MailType = Resources.MeetingRequest;
                _checkList.MailBody = string.IsNullOrEmpty(((Outlook._MeetingItem)item).Body) ? Resources.FailedToGetInformation : ((Outlook._MeetingItem)item).Body.Replace("\r\n\r\n", "\r\n");

                if (((Outlook._MeetingItem)item).RTFBody is byte[] byteArray)
                {
                    var encoding = new System.Text.ASCIIEncoding();
                    _checkList.MailHtmlBody = encoding.GetString(byteArray);
                }
                else
                {
                    _checkList.MailHtmlBody = _checkList.MailBody;
                }
            }

            _checkList.Subject = ((dynamic)item).Subject ?? Resources.FailedToGetInformation;

            _checkList = GetSenderAndSenderDomain(in item, _checkList);
            internalDomainList.Add(new InternalDomain { Domain = _checkList.SenderDomain });

            _checkList = GetAttachmentsInformation(in item, _checkList, generalSetting.IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles, attachmentsSetting, _checkList.MailHtmlBody);
            _checkList = CheckForgotAttach(_checkList, generalSetting);
            _checkList = CheckKeyword(_checkList, alertKeywordAndMessageList);
            _checkList = CheckKeywordForSubject(_checkList, alertKeywordAndMessageForSubjectList);

            var displayNameAndRecipient = MakeDisplayNameAndRecipient(((dynamic)item).Recipients, new DisplayNameAndRecipient(), generalSetting, isMailItem);

            var autoAddRecipients = AutoAddCcAndBcc(item, generalSetting, displayNameAndRecipient, autoCcBccKeywordList, autoCcBccAttachedFilesList, autoCcBccRecipientList, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, false), _checkList.Sender, generalSetting.IsAutoAddSenderToBcc, generalSetting.IsAutoAddSenderToCc);
            if (autoAddRecipients?.Count > 0)
            {
                displayNameAndRecipient = MakeDisplayNameAndRecipient(autoAddRecipients, displayNameAndRecipient, generalSetting, isMailItem);
                _ = ((dynamic)item).Recipients.ResolveAll();
            }

            displayNameAndRecipient = ExternalDomainsChangeToBccIfNeeded(item, displayNameAndRecipient, externalDomainsWarningAndAutoChangeToBccSetting, internalDomainList, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, true), _checkList.SenderDomain, _checkList.Sender, forceAutoChangeRecipientsToBccSetting);

            _checkList = GetRecipient(_checkList, displayNameAndRecipient, alertAddressList, internalDomainList);
            _checkList = CheckRecipientsAndAttachments(_checkList, attachmentsSetting.IsAttachmentsProhibited, attachmentsSetting.IsWarningWhenAttachedRealFile, attachmentProhibitedRecipientsList, recipientsAndAttachmentsNameList, attachmentAlertRecipientsList);
            _checkList = CheckMailBodyAndRecipient(_checkList, displayNameAndRecipient, nameAndDomainsList, generalSetting.IsCheckNameAndDomainsFromRecipients, generalSetting.IsCheckNameAndDomainsIncludeSubject, generalSetting.IsCheckNameAndDomainsFromSubject);
            _checkList.RecipientExternalDomainNumAll = CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, false);
            _checkList = ExternalDomainsWarningIfNeeded(_checkList, externalDomainsWarningAndAutoChangeToBccSetting, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, internalDomainList, true), forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc);
            _checkList.DeferredMinutes = CalcDeferredMinutes(displayNameAndRecipient, deferredDeliveryMinutes, generalSetting.IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain, _checkList.RecipientExternalDomainNumAll);

            if (!(contacts is null))
            {
                var contactsList = MakeContactsList(contacts);
                _checkList = AutoCheckRegisteredItemsInContacts(_checkList, displayNameAndRecipient, contactsList, generalSetting.IsAutoCheckRegisteredInContacts);
                _checkList = AddAlertOrProhibitsSendingMailIfIfRecipientsIsNotRegistered(_checkList, displayNameAndRecipient, contactsList, internalDomainList, generalSetting.IsWarningIfRecipientsIsNotRegistered, generalSetting.IsProhibitsSendingMailIfRecipientsIsNotRegistered);
            }

            return _checkList;
        }

        /// <summary>
        /// 送信元アドレスと送信元ドメインを取得。
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="checkList">CheckList</param>
        /// <returns>CheckList</returns>
        private CheckList GetSenderAndSenderDomain<T>(in T item, CheckList checkList)
        {
            try
            {
                if (typeof(T) == typeof(Outlook._MailItem) && !string.IsNullOrEmpty(((Outlook._MailItem)item).SentOnBehalfOfName))
                {
                    //代理送信の場合。
                    checkList.Sender = ((Outlook._MailItem)item).Sender?.Address ?? Resources.FailedToGetInformation;

                    if (IsValidEmailAddress(checkList.Sender))
                    {
                        //メールアドレスが取得できる場合はそのまま使う。
                        checkList.SenderDomain = checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                        checkList.Sender = $@"{checkList.Sender} ([{((Outlook._MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                    }
                    else
                    {
                        //代理送信の場合かつExchangeのCN。
                        checkList.Sender = $@"[{((Outlook._MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf}";
                        checkList.SenderDomain = @"------------------";

                        Outlook.ExchangeDistributionList exchangeDistributionList = null;
                        Outlook.ExchangeUser exchangeUser = null;

                        var sender = ((Outlook._MailItem)item).Sender;

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
                            checkList.Sender = $@"{exchangeUser.PrimarySmtpAddress} ([{((Outlook._MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                            checkList.SenderDomain = exchangeUser.PrimarySmtpAddress.Substring(exchangeUser.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                        }

                        if (!(exchangeDistributionList is null))
                        {
                            //配布リストの代理送信。
                            checkList.Sender = $@"{exchangeDistributionList.PrimarySmtpAddress} ([{((Outlook._MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                            checkList.SenderDomain = exchangeDistributionList.PrimarySmtpAddress.Substring(exchangeDistributionList.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                        }
                    }
                }
                else
                {
                    checkList.Sender = ((dynamic)item).SendUsingAccount?.SmtpAddress ?? Resources.FailedToGetInformation;

                    if (((dynamic)item).SenderEmailType == "EX" && !IsValidEmailAddress(checkList.Sender))
                    {
                        var tempOutlookApp = new Outlook.Application();
                        var tempRecipient = tempOutlookApp.Session.CreateRecipient(((dynamic)item).SenderEmailAddress);

                        try
                        {
                            _ = tempRecipient.Resolve();
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
                            checkList.Sender = ((dynamic)item).SenderEmailAddress ?? Resources.FailedToGetInformation;
                        }
                    }

                    if (!IsValidEmailAddress(checkList.Sender))
                    {
                        checkList.Sender = Resources.FailedToGetInformation;
                    }

                    checkList.SenderDomain = checkList.Sender == Resources.FailedToGetInformation ? "------------------" : checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                }
            }
            catch (Exception)
            {
                checkList.Sender = Resources.FailedToGetInformation;
                checkList.SenderDomain = @"------------------";
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
                    _ = domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }

                foreach (var recipient in displayNameAndRecipient.Cc.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    _ = domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
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

            foreach (var _ in internalDomain.Where(internalDomainSetting => domainList.Any(domain => domain.EndsWith(internalDomainSetting.Domain)) && !senderDomain.EndsWith(internalDomainSetting.Domain)))
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
            _failedToGetInformationOfRecipientsMailAddressCounter++;

            var mailAddress = Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;
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
                            mailAddress = propertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;

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
                    _ = recipient.Resolve();
                    var propertyAccessor = tempRecipient.AddressEntry.PropertyAccessor;
                    Thread.Sleep(20);

                    var isDone = false;
                    var errorCount = 0;
                    while (!isDone && errorCount < 100)
                    {
                        try
                        {
                            mailAddress = propertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;
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
        /// <param name="exchangeDistributionListMembersAreWhite">配布リストで展開したアドレスを許可リスト登録扱いするか否かの設定</param>
        /// <returns>宛先メールアドレスと宛先名称</returns>
        private IEnumerable<NameAndRecipient> GetExchangeDistributionListMembers(Outlook.Recipient recipient, bool enableGetExchangeDistributionListMembers, bool exchangeDistributionListMembersAreWhite)
        {
            _failedToGetInformationOfRecipientsMailAddressCounter++;
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
                    exchangeDistributionListMembers.Add(new NameAndRecipient { MailAddress = distributionList.PrimarySmtpAddress ?? Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter, NameAndMailAddress = (distributionList.Name ?? Resources.FailedToGetInformation) + $@" ({distributionList.PrimarySmtpAddress ?? Resources.DistributionList})" });

                    return exchangeDistributionListMembers;
                }

                var externalRecipientCounter = 1;
                var tempOutlookApp = new Outlook.Application();
                foreach (Outlook.AddressEntry member in addressEntries)
                {
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(member.Address);
                    var mailAddress = Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;

                    try
                    {
                        _ = tempRecipient.Resolve();
                        var propertyAccessor = tempRecipient.AddressEntry.PropertyAccessor;
                        Thread.Sleep(20);

                        isDone = false;
                        errorCount = 0;
                        while (!isDone && errorCount < 100)
                        {
                            try
                            {
                                mailAddress = propertyAccessor.GetProperty(@"http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString() ?? Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;
                                isDone = true;
                            }
                            catch (COMException e)
                            {
                                switch (e.ErrorCode)
                                {
                                    case -2147467260:
                                        //HRESULT:0x80004004 対策
                                        Thread.Sleep(10);
                                        errorCount++;
                                        break;
                                    case -2147467259:
                                        mailAddress = Resources.ExternalRecipient + "_" + externalRecipientCounter;
                                        externalRecipientCounter++;
                                        isDone = true;
                                        break;
                                    default:
                                        isDone = true;
                                        break;
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
        /// <param name="contactGroupMembersAreWhite">連絡先グループで展開したアドレスを許可リスト登録扱いするか否かの設定</param>
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
        /// <param name="isMailItem">メールアイテムか否か</param>
        /// <returns>宛先アドレスと名称</returns>
        private DisplayNameAndRecipient MakeDisplayNameAndRecipient(IEnumerable recipients, DisplayNameAndRecipient displayNameAndRecipient, GeneralSetting generalSetting, bool isMailItem)
        {
            foreach (Outlook.Recipient recipient in recipients)
            {
                var recipientAddressEntryUserType = Outlook.OlAddressEntryUserType.olOtherAddressEntry;
                try
                {
                    if (!isMailItem)
                    {
                        if (!recipient.Sendable) continue;
                    }

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

            if (checkList.MailBody.ToLower().Contains(Resources.AttachmentsKeyword))
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
                if (!checkList.MailBody.Contains(alertKeywordAndMessage.AlertKeyword) && alertKeywordAndMessage.AlertKeyword != "*") continue;

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
        /// 件名に登録したキーワードがある場合、登録した警告文を表示する。
        /// </summary>
        /// <param name="checkList">CheckList</param>>
        /// <param name="alertKeywordAndMessageForSubjectList">警告キーワード設定</param>>
        /// <returns>CheckList</returns>
        private CheckList CheckKeywordForSubject(CheckList checkList, IReadOnlyCollection<AlertKeywordAndMessageForSubject> alertKeywordAndMessageForSubjectList)
        {
            if (alertKeywordAndMessageForSubjectList.Count == 0) return checkList;

            foreach (var alertKeywordAndMessage in alertKeywordAndMessageForSubjectList)
            {
                if (!checkList.Subject.Contains(alertKeywordAndMessage.AlertKeyword) && alertKeywordAndMessage.AlertKeyword != "*") continue;

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
        /// <param name="item">Item</param>
        /// <param name="generalSetting">一般設定</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称設定</param>
        /// <param name="autoCcBccKeywordList">自動Cc/Bcc追加(キーワード)設定</param>
        /// <param name="autoCcBccAttachedFilesList">自動Cc/Bcc追加(キーワード)設定</param>
        /// <param name="autoCcBccRecipientList">自動Cc/Bcc追加(宛先)設定</param>
        /// <param name="externalDomainCount">外部ドメイン数</param>
        /// <param name="sender">CheckListのセンダー情報</param>
        /// <param name="isAutoAddSenderToBcc">自動で送信元アドレスをBccに追加するか否か</param>
        /// <param name="isAutoAddSenderToCc">自動で送信元アドレスをCcに追加するか否か</param>
        /// <returns>CcやBccに自動追加された宛先アドレス</returns>
        private List<Outlook.Recipient> AutoAddCcAndBcc<T>(T item, GeneralSetting generalSetting, DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<AutoCcBccKeyword> autoCcBccKeywordList, IReadOnlyCollection<AutoCcBccAttachedFile> autoCcBccAttachedFilesList, IReadOnlyCollection<AutoCcBccRecipient> autoCcBccRecipientList, int externalDomainCount, string sender, bool isAutoAddSenderToBcc, bool isAutoAddSenderToCc)
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
                            var recipient = ((dynamic)item).Recipients.Add(autoCcBccKeyword.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedCcAddressList.Add(autoCcBccKeyword.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(autoCcBccKeyword.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccKeyword.AutoAddAddress))
                    {
                        var recipient = ((dynamic)item).Recipients.Add(autoCcBccKeyword.AutoAddAddress);
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
                                var recipient = ((dynamic)item).Recipients.Add(autoCcBccAttachedFile.AutoAddAddress);
                                recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                                autoAddRecipients.Add(recipient);
                                autoAddedCcAddressList.Add(autoCcBccAttachedFile.AutoAddAddress);
                            }
                        }
                        else if (!autoAddedBccAddressList.Contains(autoCcBccAttachedFile.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccAttachedFile.AutoAddAddress))
                        {
                            var recipient = ((dynamic)item).Recipients.Add(autoCcBccAttachedFile.AutoAddAddress);
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
                            var recipient = ((dynamic)item).Recipients.Add(autoCcBccRecipient.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedCcAddressList.Add(autoCcBccRecipient.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(autoCcBccRecipient.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccRecipient.AutoAddAddress))
                    {
                        var recipient = ((dynamic)item).Recipients.Add(autoCcBccRecipient.AutoAddAddress);
                        recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recipient);
                        autoAddedBccAddressList.Add(autoCcBccRecipient.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccRecipient.CcOrBcc}] [{autoCcBccRecipient.AutoAddAddress}] (" + Resources.ApplicableDestination + $" 「{autoCcBccRecipient.TargetRecipient}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    _whitelist.Add(new Whitelist { WhiteName = autoCcBccRecipient.AutoAddAddress });
                }
            }

            //常に自分をCcまたはBccに入れるオプションが有効な場合、それをする。
            if (isAutoAddSenderToCc || isAutoAddSenderToBcc)
            {
                var addSenderToCc = isAutoAddSenderToCc;
                var addSenderToBcc = isAutoAddSenderToBcc;

                var mailItemSender = ((dynamic)item).SenderEmailAddress;

                if (typeof(T) == typeof(Outlook._MailItem))
                {
                    if (!string.IsNullOrEmpty(((Outlook._MailItem)item).SentOnBehalfOfName) && !string.IsNullOrEmpty(((Outlook._MailItem)item).Sender.Address))
                    {
                        mailItemSender = ((Outlook._MailItem)item).Sender.Address;
                    }
                }

                var counter = 0;
                while (counter <= 5)
                {
                    counter++;
                    try
                    {
                        foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
                        {
                            switch (recipient.Type)
                            {
                                case (int)Outlook.OlMailRecipientType.olBCC when recipient.Address.Equals(mailItemSender):
                                    addSenderToBcc = false;
                                    break;
                                case (int)Outlook.OlMailRecipientType.olCC when recipient.Address.Equals(mailItemSender):
                                    addSenderToCc = false;
                                    break;
                            }
                        }
                        counter = 6;
                        break;
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(10);
                    }
                }

                if (addSenderToCc)
                {
                    counter = 0;
                    while (counter <= 3)
                    {
                        counter++;
                        try
                        {
                            var senderAsRecipient = ((dynamic)item).Recipients.Add(mailItemSender);
                            Thread.Sleep(150);

                            _ = senderAsRecipient.Resolve();
                            Thread.Sleep(150);

                            senderAsRecipient.Type = (int)Outlook.OlMailRecipientType.olCC;
                            autoAddRecipients.Add(senderAsRecipient);
                            mailItemSender = senderAsRecipient.Address;
                            counter = 4;
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(10);
                        }
                    }
                }

                if (addSenderToBcc)
                {
                    counter = 0;
                    while (counter < 3)
                    {
                        counter++;
                        try
                        {
                            var senderAsRecipient = ((dynamic)item).Recipients.Add(mailItemSender);
                            Thread.Sleep(150);

                            _ = senderAsRecipient.Resolve();
                            Thread.Sleep(150);

                            senderAsRecipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                            autoAddRecipients.Add(senderAsRecipient);
                            mailItemSender = senderAsRecipient.Address;
                            counter = 4;
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(10);
                        }
                    }
                }

                _whitelist.Add(new Whitelist { WhiteName = sender, IsSkipConfirmation = false });
            }

            return autoAddRecipients;
        }

        /// <summary>
        /// HTML内に埋め込まれた添付ファイル名を取得する。
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="mailHtmlBody">メール本文(HTML形式)</param>
        /// <returns>埋め込みファイル名のList</returns>
        private List<string> MakeEmbeddedAttachmentsList<T>(T item, string mailHtmlBody)
        {
            if (typeof(T) == typeof(Outlook._MailItem))
            {
                //HTML形式の場合のみ、処理対象とする。
                if (((Outlook._MailItem)item).BodyFormat != Outlook.OlBodyFormat.olFormatHTML) return null;
            }

            var matches = Regex.Matches(mailHtmlBody, @"cid:.*?@");

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
        /// <param name="item">Item</param>
        /// <param name="checkList">CheckList</param>
        /// <param name="isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles">HTML埋め込みの添付ファイル無視設定</param>
        /// <param name="attachmentsSetting">添付ファイルに関する設定</param>
        /// <param name="mailHtmlBody">メール本文(HTML形式)</param>
        /// <returns>CheckList</returns>
        private CheckList GetAttachmentsInformation<T>(in T item, CheckList checkList, bool isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles, AttachmentsSetting attachmentsSetting, string mailHtmlBody)
        {
            if (((dynamic)item).Attachments.Count == 0) return checkList;

            var embeddedAttachmentsName = new List<string>();
            if (isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles)
            {
                embeddedAttachmentsName = MakeEmbeddedAttachmentsList(item, mailHtmlBody);
            }

            var tempDirectoryPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(tempDirectoryPath);

            for (var i = 0; i < ((dynamic)item).Attachments.Count; i++)
            {
                var fileSize = "?KB";
                if (((dynamic)item).Attachments[i + 1].Size != 0)
                {
                    fileSize = Math.Round(((double)((dynamic)item).Attachments[i + 1].Size / 1024), 0, MidpointRounding.AwayFromZero).ToString("##,###") + "KB";
                }

                if (((dynamic)item).Attachments[i + 1].Size >= 10485760)
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsBigAttachedFile + $"[{((dynamic)item).Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                }

                //一部の状態で添付ファイルのファイルタイプを取得できないため、それを回避。
                string fileType;
                try
                {
                    fileType = ((dynamic)item).Attachments[i + 1].FileName.Substring(((dynamic)item).Attachments[i + 1].FileName.LastIndexOf(".", StringComparison.Ordinal));
                }
                catch (Exception)
                {
                    fileType = Resources.Unknown;
                }

                var isDangerous = false;
                if (fileType == ".exe")
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsAttachedExe + $"[{((dynamic)item).Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                    isDangerous = true;
                }

                string fileName;
                try
                {
                    fileName = ((dynamic)item).Attachments[i + 1].FileName;
                }
                catch (Exception)
                {
                    fileName = Resources.Unknown;
                }

                //情報取得に完全に失敗した添付ファイルは無視する。(リッチテキスト形式の埋め込み画像など)
                if (fileName == Resources.Unknown && fileSize == "?KB" && fileType == Resources.Unknown) continue;

                //電子署名付きメールの証明書はあえて無視する。
                if (fileType == ".p7s" || fileType == "p7s") continue;

                var isEncrypted = false;

                try
                {
                    if ((attachmentsSetting.IsWarningWhenEncryptedZipIsAttached || attachmentsSetting.IsProhibitedWhenEncryptedZipIsAttached) && fileName != Resources.Unknown)
                    {
                        if (attachmentsSetting.IsEnableAllAttachedFilesAreDetectEncryptedZip || fileType == ".zip" || fileType == "zip")
                        {
                            var tempFilePath = Path.Combine(tempDirectoryPath, fileName);
                            ((dynamic)item).Attachments[i + 1].SaveAsFile(tempFilePath);

                            var zipTools = new ZipTools();
                            if (zipTools.CheckZipIsEncrypted(tempFilePath))
                            {
                                File.Delete(tempFilePath);

                                isEncrypted = true;
                                AddAlerts(Resources.AttachedIsAnEncryptedZipFile + $" [{fileName}]", true, false, false);

                                if (attachmentsSetting.IsProhibitedWhenEncryptedZipIsAttached)
                                {
                                    checkList.IsCanNotSendMail = true;
                                    checkList.CanNotSendMailMessage = Resources.AttachedIsAnEncryptedZipFile + $"{Environment.NewLine}[{fileName}]";
                                }
                            }

                            File.Delete(tempFilePath);
                        }
                    }
                }
                catch (Exception)
                {
                    //Do Nothing.
                }

                if (embeddedAttachmentsName is null)
                {
                    checkList.Attachments.Add(new Attachment
                    {
                        FileName = fileName,
                        FileSize = fileSize,
                        FileType = fileType,
                        IsTooBig = ((dynamic)item).Attachments[i + 1].Size >= 10485760,
                        IsEncrypted = isEncrypted,
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
                        IsTooBig = ((dynamic)item).Attachments[i + 1].Size >= 10485760,
                        IsEncrypted = isEncrypted,
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
        /// <param name="isCheckNameAndDomainsFromRecipients">本文内に宛先名称が無い場合にも警告を表示するか否か</param>
        /// <param name="isCheckNameAndDomainsIncludeSubject">対象に件名を含めるか否か</param>
        /// <param name="isCheckNameAndDomainsFromSubject">件名内に宛先名称が無い場合にも警告を表示するか否か</param>
        /// <returns>CheckList</returns>
        private CheckList CheckMailBodyAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<NameAndDomains> nameAndDomainsList, bool isCheckNameAndDomainsFromRecipients, bool isCheckNameAndDomainsIncludeSubject, bool isCheckNameAndDomainsFromSubject)
        {
            if (displayNameAndRecipient is null) return checkList;

            //空の設定値があると誤検知するため、空を省く。
            var cleanedNameAndDomains = nameAndDomainsList.Where(nameAndDomain => !string.IsNullOrEmpty(nameAndDomain.Domain) && !string.IsNullOrEmpty(nameAndDomain.Name)).ToList();

            if (isCheckNameAndDomainsFromRecipients || (isCheckNameAndDomainsIncludeSubject && isCheckNameAndDomainsFromSubject))
            {
                var domainCandidateRecipients = new List<string[]>();
                foreach (var nameAndDomain in cleanedNameAndDomains)
                {
                    foreach (var recipient in displayNameAndRecipient.All)
                    {
                        if (recipient.Value.Contains(nameAndDomain.Domain))
                        {
                            domainCandidateRecipients.Add(new[] { recipient.Value, nameAndDomain.Name });
                        }
                    }
                }

                if (isCheckNameAndDomainsFromRecipients)
                {
                    foreach (var domainAndName in domainCandidateRecipients.Where(domainAndName => !checkList.MailBody.Contains(domainAndName[1])).Where(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain)))
                    {
                        checkList.Alerts.Add(new Alert
                        {
                            AlertMessage = $"{domainAndName[0]} : {Resources.CanNotFindTheLinkedName} ({domainAndName[1]})",
                            IsImportant = true,
                            IsWhite = false,
                            IsChecked = false
                        });
                    }
                }

                if (isCheckNameAndDomainsIncludeSubject && isCheckNameAndDomainsFromSubject)
                {
                    foreach (var domainAndName in domainCandidateRecipients.Where(domainAndName => !checkList.Subject.Contains(domainAndName[1])).Where(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain)))
                    {
                        checkList.Alerts.Add(new Alert
                        {
                            AlertMessage = $"{domainAndName[0]} : {Resources.CanNotFindTheLinkedNameInSubject} ({domainAndName[1]})",
                            IsImportant = true,
                            IsWhite = false,
                            IsChecked = false
                        });
                    }
                }
            }

            var targetText = checkList.MailBody;
            if (isCheckNameAndDomainsIncludeSubject) { targetText += checkList.Subject; }

            var recipientCandidateDomains = (from nameAndDomain in cleanedNameAndDomains where targetText.Contains(nameAndDomain.Name) select nameAndDomain.Domain).ToList();
            //送信先の候補が見つからない場合、これ以上何もしない。(見つからない場合の方が多いため、警告ばかりになってしまう。)
            if (recipientCandidateDomains.Count == 0) return checkList;

            foreach (var recipient in displayNameAndRecipient.All)
            {
                if (recipientCandidateDomains.Any(domains => domains.Equals(recipient.Key.Substring(recipient.Key.IndexOf("@", StringComparison.Ordinal))))) continue;

                //送信者ドメインは警告対象外。
                if (!recipient.Key.Contains(checkList.SenderDomain))
                {
                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = recipient.Value + " : " + Resources.IsAlertAddressMaybeIrrelevant,
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
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
        private CheckList GetRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<AlertAddress> alertAddressList, IReadOnlyCollection<InternalDomain> internalDomainList)
        {
            if (displayNameAndRecipient is null) return checkList;

            foreach (var to in displayNameAndRecipient.To)
            {
                var isExternal = true;
                foreach (var _ in internalDomainList.Where(internalDomainSetting => to.Key.EndsWith(internalDomainSetting.Domain)))
                {
                    isExternal = false;
                }

                if (to.Value.Contains(Resources.DistributionList) && to.Key.Contains(Resources.FailedToGetInformation))
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

                if (alertAddressList.Count == 0) continue;

                foreach (var alertAddress in alertAddressList)
                {
                    if (!to.Key.Contains(alertAddress.TargetAddress)) continue;

                    if (alertAddress.IsCanNotSend)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{to.Value}]";
                        continue;
                    }

                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{to.Value}]" : alertAddress.Message + $"[{to.Value}]",
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isExternal = true;
                foreach (var _ in internalDomainList.Where(internalDomainSetting => cc.Key.EndsWith(internalDomainSetting.Domain)))
                {
                    isExternal = false;
                }

                if (cc.Value.Contains(Resources.DistributionList) && cc.Key.Contains(Resources.FailedToGetInformation))
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

                if (alertAddressList.Count == 0) continue;

                foreach (var alertAddress in alertAddressList)
                {
                    if (!cc.Key.Contains(alertAddress.TargetAddress)) continue;

                    if (alertAddress.IsCanNotSend)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{cc.Value}]";
                        continue;
                    }

                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{cc.Value}]" : alertAddress.Message + $"[{cc.Value}]",
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            foreach (var bcc in displayNameAndRecipient.Bcc)
            {
                var isExternal = true;
                foreach (var _ in internalDomainList.Where(internalDomainSetting => bcc.Key.EndsWith(internalDomainSetting.Domain)))
                {
                    isExternal = false;
                }

                if (bcc.Value.Contains(Resources.DistributionList) && bcc.Key.Contains(Resources.FailedToGetInformation))
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

                if (alertAddressList.Count == 0) continue;

                foreach (var alertAddress in alertAddressList)
                {
                    if (!bcc.Key.Contains(alertAddress.TargetAddress)) continue;

                    if (alertAddress.IsCanNotSend)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{bcc.Value}]";
                        continue;
                    }

                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{bcc.Value}]" : alertAddress.Message + $"[{bcc.Value}]",
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
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
        /// <param name="isForceAutoChangeRecipientsToBcc">強制的に全ての宛先をBccに変換するか否か</param>
        /// <returns>CheckList</returns>
        private CheckList ExternalDomainsWarningIfNeeded(CheckList checkList, ExternalDomainsWarningAndAutoChangeToBcc settings, int externalDomainNumToAndCc, bool isForceAutoChangeRecipientsToBcc)
        {
            //強制Bcc変換が有効な場合、この機能は無視する。
            if (isForceAutoChangeRecipientsToBcc) return checkList;

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
        /// <param name="item">Item</param>
        /// <param name="mailItemsRecipientAndMailAddress">メールアドレスとMailItemのRecipient</param>
        /// <param name="senderMailAddress">送信元メールアドレス</param>
        /// <param name="isNeedsAddToSender">Toへ送信元アドレスを追加するか否か</param>
        private void ChangeToBcc<T>(T item, IReadOnlyCollection<MailItemsRecipientAndMailAddress> mailItemsRecipientAndMailAddress, string senderMailAddress, bool isNeedsAddToSender)
        {
            if ((dynamic)item is null) return;

            var targetMailAddressAndRecipient = new Dictionary<string, string>();

            foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
            {
                foreach (var target in mailItemsRecipientAndMailAddress)
                {
                    if (recipient.Address == target.MailItemsRecipient) targetMailAddressAndRecipient[target.MailAddress] = target.MailItemsRecipient;
                }
            }

            //Indexを使用してRemoveした場合、Indexがずれ、複数を正しく削除できないため、削除対象を探して削除する。
            var targetCount = targetMailAddressAndRecipient.Count;
            while (targetCount > 0)
            {
                foreach (var target in targetMailAddressAndRecipient)
                {
                    foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
                    {
                        if (recipient.Address != target.Value) continue;
                        ((dynamic)item).Recipients.Remove(recipient.Index);
                        targetCount--;
                    }
                }
            }

            foreach (var addTarget in targetMailAddressAndRecipient.Select(mailAddress => ((dynamic)item).Recipients.Add(mailAddress.Key)))
            {
                addTarget.Type = (int)Outlook.OlMailRecipientType.olBCC;
            }

            if (isNeedsAddToSender)
            {
                var senderRecipient = ((dynamic)item).Recipients.Add(senderMailAddress);
                senderRecipient.Type = (int)Outlook.OlMailRecipientType.olTo;
            }

            _ = ((dynamic)item).Recipients.ResolveAll();
        }

        /// <summary>
        /// 条件に該当する場合、ToとCcの外部アドレスをBccに変換する。
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="settings">外部ドメイン数の警告と自動Bcc変換の設定</param>
        /// <param name="internalDomains">内部ドメイン</param>
        /// <param name="externalDomainNumToAndCc">ToとCcの外部ドメイン数</param>
        /// <param name="senderDomain">送信元ドメイン</param>
        /// <param name="senderMailAddress">送信元アドレス</param>
        /// <param name="forceAutoChangeRecipientsToBccSetting">forceAutoChangeRecipientsToBccSetting</param>
        /// <returns>DisplayNameAndRecipient</returns>
        private DisplayNameAndRecipient ExternalDomainsChangeToBccIfNeeded<T>(T item, DisplayNameAndRecipient displayNameAndRecipient, ExternalDomainsWarningAndAutoChangeToBcc settings, ICollection<InternalDomain> internalDomains, int externalDomainNumToAndCc, string senderDomain, string senderMailAddress, ForceAutoChangeRecipientsToBcc forceAutoChangeRecipientsToBccSetting)
        {
            if ((!settings.IsAutoChangeToBccWhenLargeNumberOfExternalDomains || settings.IsProhibitedWhenLargeNumberOfExternalDomains || settings.TargetToAndCcExternalDomainsNum > externalDomainNumToAndCc) && !forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc) return displayNameAndRecipient;

            if (forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc && forceAutoChangeRecipientsToBccSetting.IsIncludeInternalDomain)
            {
                internalDomains.Clear();
            }
            else
            {
                internalDomains.Add(new InternalDomain { Domain = senderDomain });
            }

            var removeTarget = new List<string>();

            foreach (var to in displayNameAndRecipient.To)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => to.Key.EndsWith(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[to.Key] = to.Value;
                removeTarget.Add(to.Key);
            }
            foreach (var target in removeTarget)
            {
                _ = displayNameAndRecipient.To.Remove(target);
            }

            removeTarget.Clear();

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => cc.Key.EndsWith(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[cc.Key] = cc.Value;
                removeTarget.Add(cc.Key);
            }
            foreach (var target in removeTarget)
            {
                _ = displayNameAndRecipient.Cc.Remove(target);
            }

            if (forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc)
            {
                AddAlerts(Resources.ForceAutoChangeRecipientsToBccAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]", false, false, true);
            }
            else
            {
                AddAlerts(Resources.ExternalDomainsChangeToBccAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]", true, false, false);
            }
            //Toが存在しなくなる場合、送信者をToに追加する。
            var isNeedsAddToSender = false;
            var thisSenderMailAddress = forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc && !string.IsNullOrEmpty(forceAutoChangeRecipientsToBccSetting.ToRecipient) ? forceAutoChangeRecipientsToBccSetting.ToRecipient : senderMailAddress;
            if (displayNameAndRecipient.To.Count == 0)
            {
                displayNameAndRecipient.To[thisSenderMailAddress] = thisSenderMailAddress;
                isNeedsAddToSender = true;

                AddAlerts(thisSenderMailAddress == senderMailAddress
                        ? Resources.AutoAddSendersAddressToAlert
                        : Resources.AutoAddToRecipientByForceAutoChangeRecipientsToBccAddressToAlert, true, false, false);
            }

            var targetMailRecipientsIndex = new List<MailItemsRecipientAndMailAddress>();
            //元からBccの宛先には何もしない。
            foreach (var recipient in displayNameAndRecipient.MailRecipientsIndex.Where(recipient => recipient.Type != (int)Outlook.OlMailRecipientType.olBCC))
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(internalDomain => recipient.MailAddress.EndsWith(internalDomain.Domain)))
                {
                    isExternal = false;
                }

                if (isExternal) targetMailRecipientsIndex.Add(recipient);
            }

            ChangeToBcc(item, targetMailRecipientsIndex, thisSenderMailAddress, isNeedsAddToSender);

            return displayNameAndRecipient;
        }

        /// <summary>
        /// 宛先と添付ファイル名の紐づけなどを確認する。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="isAttachmentsProhibited">添付ファイル付きのメールを送信を禁止するか否か</param>
        /// <param name="isWarningWhenAttachedRealFile">実ファイルが添付されている場合、リンクとして添付を推奨する旨の警告を表示する否か</param>
        /// <param name="attachmentProhibitedRecipientsList">添付ファイル禁止宛先設定</param>
        /// <param name="recipientsAndAttachmentsNameList">宛先と添付ファイル名の紐づけ設定</param>
        /// <param name="attachmentAlertRecipientsList">添付ファイル警告宛先と警告文設定</param>
        /// <returns>CheckList</returns>
        private CheckList CheckRecipientsAndAttachments(CheckList checkList, bool isAttachmentsProhibited, bool isWarningWhenAttachedRealFile, IReadOnlyCollection<AttachmentProhibitedRecipients> attachmentProhibitedRecipientsList, IReadOnlyCollection<RecipientsAndAttachmentsName> recipientsAndAttachmentsNameList, IReadOnlyCollection<AttachmentAlertRecipients> attachmentAlertRecipientsList)
        {
            if (checkList.Attachments.Count <= 0) return checkList;

            if (isAttachmentsProhibited)
            {
                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = Resources.AttachmentsProhibitedMessage;

                //添付ファイル付きメールの送信が禁止されているため、これ以上何もしない。
                return checkList;
            }

            if (attachmentProhibitedRecipientsList.Count > 0)
            {
                var prohibitedRecipients = "";
                var isProhibited = false;

                foreach (var prohibitedRecipient in attachmentProhibitedRecipientsList)
                {
                    foreach (var to in checkList.ToAddresses.Where(to => to.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + to.MailAddress;
                    }

                    foreach (var cc in checkList.CcAddresses.Where(cc => cc.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + cc.MailAddress;
                    }

                    foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + bcc.MailAddress;
                    }
                }

                if (isProhibited)
                {
                    checkList.CanNotSendMailMessage = Resources.AttachmentProhibitedRecipientsMessage + "：" + prohibitedRecipients;

                    //添付ファイル付きメールの送付が禁止された宛先のため、これ以上何もしない。
                    return checkList;
                }
            }

            if (attachmentAlertRecipientsList.Count > 0)
            {
                foreach (var attachmentAlertRecipient in attachmentAlertRecipientsList)
                {
                    foreach (var to in checkList.ToAddresses.Where(to => to.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{to.MailAddress}]" : attachmentAlertRecipient.Message + $"[{to.MailAddress}]", true, false, false);
                    }
                    foreach (var cc in checkList.CcAddresses.Where(cc => cc.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{cc.MailAddress}]" : attachmentAlertRecipient.Message + $"[{cc.MailAddress}]", true, false, false);
                    }
                    foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{bcc.MailAddress}]" : attachmentAlertRecipient.Message + $"[{bcc.MailAddress}]", true, false, false);
                    }
                }
            }

            if (recipientsAndAttachmentsNameList.Count > 0)
            {
                foreach (var recipientsAndAttachmentsName in recipientsAndAttachmentsNameList)
                {
                    foreach (var attachment in checkList.Attachments.Where(attachment => attachment.FileName.Contains(recipientsAndAttachmentsName.AttachmentsName)))
                    {
                        foreach (var to in checkList.ToAddresses.Where(to => to.IsExternal))
                        {
                            if (!to.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(Resources.RecipientsAndAttachmentsNameMessage + "：" + to.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }

                        foreach (var cc in checkList.CcAddresses.Where(cc => cc.IsExternal))
                        {
                            if (!cc.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(Resources.RecipientsAndAttachmentsNameMessage + "：" + cc.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }

                        foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.IsExternal))
                        {
                            if (!bcc.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(Resources.RecipientsAndAttachmentsNameMessage + "：" + bcc.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }
                    }
                }
            }

            if (isWarningWhenAttachedRealFile)
            {
                AddAlerts(Resources.RecommendationOfAttachFileAsLink, false, true, false);
            }

            return checkList;
        }

        /// <summary>
        /// 連絡先(アドレス帳)に登録された宛先にあらかじめチェックする。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="contactsList">連絡先(アドレス帳)</param>
        /// <param name="isAutoCheckRegisteredInContacts">連絡先(アドレス帳)に登録された宛先にあらかじめチェックを入れるか否か</param>
        /// <returns>CheckList</returns>
        private CheckList AutoCheckRegisteredItemsInContacts(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<string> contactsList, bool isAutoCheckRegisteredInContacts)
        {
            if (!isAutoCheckRegisteredInContacts) return checkList;

            foreach (var mailItemsRecipient in contactsList.SelectMany(contact => displayNameAndRecipient.MailRecipientsIndex.Where(mailItemsRecipient => contact == mailItemsRecipient.MailAddress || contact == mailItemsRecipient.MailItemsRecipient)))
            {
                foreach (var toAddress in checkList.ToAddresses.Where(toAddress => toAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    toAddress.IsChecked = true;
                }

                foreach (var ccAddress in checkList.CcAddresses.Where(ccAddress => ccAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    ccAddress.IsChecked = true;
                }

                foreach (var bccAddress in checkList.BccAddresses.Where(bccAddress => bccAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    bccAddress.IsChecked = true;
                }
            }

            return checkList;
        }

        /// <summary>
        /// 連絡先(アドレス帳)未登録の宛先に対し、警告を表示したり送信をブロックしたりする。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">宛先アドレスと名称</param>
        /// <param name="contactsList">連絡先(アドレス帳)</param>
        /// <param name="internalDomainList">内部ドメイン</param>
        /// <param name="isWarningIfRecipientsIsNotRegistered">連絡先(アドレス帳)未登録の宛先に警告を表示するか否か</param>
        /// <param name="isProhibitsSendingMailIfRecipientsIsNotRegistered">連絡先(アドレス帳)未登録の宛先が存在する場合、メールの送信を禁止するか否か</param>
        /// <returns>CheckList</returns>
        private CheckList AddAlertOrProhibitsSendingMailIfIfRecipientsIsNotRegistered(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<string> contactsList, IReadOnlyCollection<InternalDomain> internalDomainList, bool isWarningIfRecipientsIsNotRegistered, bool isProhibitsSendingMailIfRecipientsIsNotRegistered)
        {
            if (!(isWarningIfRecipientsIsNotRegistered || isProhibitsSendingMailIfRecipientsIsNotRegistered)) return checkList;

            var selectedContactsList = contactsList.SelectMany(contact => displayNameAndRecipient.MailRecipientsIndex.Where(mailItemsRecipient => contact == mailItemsRecipient.MailAddress || contact == mailItemsRecipient.MailItemsRecipient)).Select(x => x.MailAddress).ToList();

            foreach (var to in displayNameAndRecipient.To.Where(to => !selectedContactsList.Contains(to.Key)))
            {
                //内部ドメインは対象外
                if (internalDomainList.Any(internalDomain => to.Key.EndsWith(internalDomain.Domain))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{to.Value}]";
                    return checkList;
                }

                AddAlerts(Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{to.Value}]", true, false, false);
            }

            foreach (var cc in displayNameAndRecipient.Cc.Where(cc => !selectedContactsList.Contains(cc.Key)))
            {
                //内部ドメインは対象外
                if (internalDomainList.Any(internalDomain => cc.Key.EndsWith(internalDomain.Domain))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{cc.Value}]";
                    return checkList;
                }

                AddAlerts(Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{cc.Value}]", true, false, false);
            }

            foreach (var bcc in displayNameAndRecipient.Bcc.Where(bcc => !selectedContactsList.Contains(bcc.Key)))
            {
                //内部ドメインは対象外
                if (internalDomainList.Any(internalDomain => bcc.Key.EndsWith(internalDomain.Domain))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{bcc.Value}]";
                    return checkList;
                }

                AddAlerts(Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{bcc.Value}]", true, false, false);
            }

            return checkList;
        }

        /// <summary>
        /// 本文のプレビューに文言を追加する。(この時点で、実際のメール文面には追加しない。※送信をキャンセルする可能性があるため)
        /// </summary>
        /// <param name="mailBody">メール本文(テキスト形式)</param>
        /// <param name="autoAddMessageSetting">autoAddMessageSetting</param>
        /// <returns>メール本文(テキスト形式)</returns>
        private string AddMessageToBodyPreview(string mailBody, AutoAddMessage autoAddMessageSetting)
        {
            if (mailBody == Resources.FailedToGetInformation) return mailBody;

            if (autoAddMessageSetting.IsAddToStart)
            {
                mailBody = autoAddMessageSetting.MessageOfAddToStart + Environment.NewLine + Environment.NewLine + mailBody;
                AddAlerts(Resources.AddedTextAtTheBeginning, false, false, true);
            }

            if (autoAddMessageSetting.IsAddToEnd)
            {
                mailBody = mailBody + Environment.NewLine + autoAddMessageSetting.MessageOfAddToEnd;
                AddAlerts(Resources.AddedTextAtTheEnd, false, false, true);
            }

            return mailBody;
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

        #region Tools

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

        /// <summary>
        /// 連絡先に登録されているアドレスをすべて取得する。
        /// </summary>
        /// <param name="contacts">連絡先フォルダ</param>
        /// <returns>連絡先に登録されているアドレスのリスト</returns>
        private List<string> MakeContactsList(Outlook.MAPIFolder contacts)
        {
            if (contacts is null) return null;

            var contactsList = new List<string>();
            foreach (dynamic contact in contacts.Items)
            {
                if (!(contact is Outlook.ContactItem foundContact))
                {
                    try
                    {
                        var entryId = contact.EntryID;

                        var tempOutlookApp = new Outlook.Application().GetNamespace("MAPI");
                        var distList = (Outlook.DistListItem)tempOutlookApp.GetItemFromID(entryId);

                        for (var i = 1; i < distList.MemberCount + 1; i++)
                        {
                            if (!(distList.GetMember(i).Address is null) && distList.GetMember(i).Address != "Unknown")
                            {
                                contactsList.Add(distList.GetMember(i).Address);
                            }

                        }
                    }
                    catch (Exception)
                    {
                        // Do Nothing.
                    }
                }
                else
                {
                    if (!(foundContact.Email1Address is null))
                    {
                        contactsList.Add(foundContact.Email1Address);
                        if (IsValidEmailAddress(foundContact.Email1Address)) continue;
                        //登録アドレスがメールアドレスでない場合、ExchangeのCN(X.500)の可能性があるため、一般的なメールアドレスに変換したものも併せて登録する。
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(foundContact.Email1Address);
                        if (!(exchangePrimarySmtpAddress is null))
                        {
                            contactsList.Add(exchangePrimarySmtpAddress);
                        }
                    }
                    else if (!(foundContact.Email2Address is null))
                    {
                        contactsList.Add(foundContact.Email2Address);
                        if (IsValidEmailAddress(foundContact.Email2Address)) continue;
                        //登録アドレスがメールアドレスでない場合、ExchangeのCN(X.500)の可能性があるため、一般的なメールアドレスに変換したものも併せて登録する。
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(foundContact.Email2Address);
                        if (!(exchangePrimarySmtpAddress is null))
                        {
                            contactsList.Add(exchangePrimarySmtpAddress);
                        }
                    }
                    else if (!(foundContact.Email3Address is null))
                    {
                        contactsList.Add(foundContact.Email3Address);
                        if (IsValidEmailAddress(foundContact.Email3Address)) continue;
                        //登録アドレスがメールアドレスでない場合、ExchangeのCN(X.500)の可能性があるため、一般的なメールアドレスに変換したものも併せて登録する。
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(foundContact.Email3Address);
                        if (!(exchangePrimarySmtpAddress is null))
                        {
                            contactsList.Add(exchangePrimarySmtpAddress);
                        }
                    }
                }

            }

            return contactsList;
        }

        /// <summary>
        /// X500形式のアドレスを一般的なメールアドレスに変換する。
        /// </summary>
        /// <param name="x500">x500形式のアドレス</param>
        /// <returns>一般的なメールアドレス</returns>
        private string GetExchangePrimarySmtpAddress(string x500)
        {
            var tempOutlookApp = new Outlook.Application();
            var tempRecipient = tempOutlookApp.Session.CreateRecipient(x500);

            try
            {
                _ = tempRecipient.Resolve();
                var addressEntry = tempRecipient.AddressEntry;

                var isDone = false;
                var errorCount = 0;
                while (!isDone && errorCount < 100)
                {
                    try
                    {
                        var exchangeUser = addressEntry?.GetExchangeUser();
                        if (exchangeUser?.PrimarySmtpAddress != null)
                        {
                            return exchangeUser.PrimarySmtpAddress;
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
            }
            catch (Exception)
            {
                //Do Nothing.
            }

            return null;
        }

        #endregion

    }
}