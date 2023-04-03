using OutlookOkan.Handlers;
using OutlookOkan.Helpers;
using OutlookOkan.Models;
using OutlookOkan.Services;
using OutlookOkan.Types;
using OutlookOkan.Views;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookOkan
{
    public partial class ThisAddIn
    {
        private readonly GeneralSetting _generalSetting = new GeneralSetting();
        private readonly SecurityForReceivedMail _securityForReceivedMail = new SecurityForReceivedMail();
        private readonly List<AlertKeywordOfSubjectWhenOpeningMail> _alertKeywordOfSubjectWhenOpeningMail = new List<AlertKeywordOfSubjectWhenOpeningMail>();

        private Outlook.Inspectors _inspectors;
        private Outlook.Explorer _currentExplorer;
        private Outlook.MailItem _currentMailItem;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        /// <summary>
        /// アドインのロード時(Outlookの起動時)の処理。
        /// </summary>
        /// <param name="sender">Sender</param>
        /// <param name="e">EventArgs</param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //ユーザ設定をロード。(このタイミングでロードしないと、リボンの表示言語を変更できない。)
            LoadGeneralSetting(true);
            if (!(_generalSetting.LanguageCode is null))
            {
                ResourceService.Instance.ChangeCulture(_generalSetting.LanguageCode);
            }

            LoadSecurityForReceivedMail();
            if (_securityForReceivedMail.IsEnableSecurityForReceivedMail)
            {
                LoadAlertKeywordOfSubjectWhenOpeningMailsData();
                _currentExplorer = Application.ActiveExplorer();
                _currentExplorer.SelectionChange += CurrentExplorer_SelectionChange;
            }

            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += OpenOutboxItemInspector;

            Application.ItemSend += Application_ItemSend;
        }

        private string _currentMailItemEntryId = "";
        private void CurrentExplorer_SelectionChange()
        {
            var currentExplorer = Application.ActiveExplorer();
            if (currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJournal).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderRssFeeds).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderServerFailures).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderLocalFailures).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSyncIssues).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks).Name
                || currentExplorer.CurrentFolder.Name == Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo).Name
               )
            {
                return;
            }

            var selection = currentExplorer.Selection;
            if (selection is null || selection.Count != 1) return;
            if (!(selection[1] is Outlook.MailItem selectedMail)) return;

            _currentMailItem = selectedMail;
            if (_currentMailItemEntryId == _currentMailItem.EntryID) return;

            _currentMailItemEntryId = _currentMailItem.EntryID;

            //件名にキーワードが含まれている場合の警告機能
            if (_securityForReceivedMail.IsEnableAlertKeywordOfSubjectWhenOpeningMailsData)
            {
                var subject = selectedMail.Subject;
                var settings = _alertKeywordOfSubjectWhenOpeningMail.FirstOrDefault(x => subject.Contains(x.AlertKeyword));

                if (!(settings is null))
                {
                    var message = Properties.Resources.AlertOfReceivedMailSubject + Environment.NewLine + "[" + settings.AlertKeyword + "]";
                    if (!string.IsNullOrEmpty(settings.Message))
                    {
                        message = settings.Message;
                    }
                    MessageBox.Show(message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }

            //メールヘッダの解析と警告機能
            if (_securityForReceivedMail.IsEnableMailHeaderAnalysis)
            {
                var header = selectedMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");
                var analysisResults = MailHeaderHandler.ValidateEmailHeader(header.ToString());
                if (!(analysisResults is null))
                {
                    var message = "";
                    foreach (KeyValuePair<string, string> entry in analysisResults)
                    {
                        message += ($"{entry.Key}: {entry.Value}") + Environment.NewLine;
                    }

                    //SPFレコードの検証に失敗した場合に警告を表示する。
                    if (_securityForReceivedMail.IsShowWarningWhenSpfFails)
                    {
                        if (analysisResults["SPF"] == "FAIL" || analysisResults["SPF"] == "NONE")
                        {

                            _ = MessageBox.Show(Properties.Resources.SpfWarning1 + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }

                    //DKIMレコードの検証に失敗した場合に警告を表示する。
                    if (_securityForReceivedMail.IsShowWarningWhenDkimFails)
                    {
                        if (analysisResults["DKIM"] == "FAIL")
                        {
                            _ = MessageBox.Show(Properties.Resources.DkimWarning1 + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
            }

            //添付ファイルを開くときの警告機能
            if (_securityForReceivedMail.IsEnableWarningFeatureWhenOpeningAttachments && selectedMail.Attachments.Count != 0)
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                Thread.Sleep(10);

                _currentMailItem.BeforeAttachmentRead -= BeforeAttachmentRead;
                _currentMailItem.BeforeAttachmentRead += BeforeAttachmentRead;
            }
        }

        /// <summary>
        /// 添付ファイルを開く時の分析と警告
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
        private void BeforeAttachmentRead(Outlook.Attachment attachment, ref bool cancel)
        {
            //添付ファイルを開く前の警告機能
            if (_securityForReceivedMail.IsWarnBeforeOpeningAttachments)
            {
                var dialogResult = MessageBox.Show(Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Properties.Resources.OpenAttachmentWarning2 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dialogResult == MessageBoxResult.Yes)
                {
                    //Open file.
                }
                else
                {
                    cancel = true;
                    return;
                }
            }

            if (_securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip || _securityForReceivedMail.IsWarnLinkFileInTheZip || _securityForReceivedMail.IsWarnOneFileInTheZip || _securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip || _securityForReceivedMail.IsWarnBeforeOpeningAttachmentsThatContainMacros)
            {
                var tempDirectoryPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
                _ = Directory.CreateDirectory(tempDirectoryPath);
                var tempFilePath = Path.Combine(tempDirectoryPath, Guid.NewGuid().ToString("N"));
                attachment.SaveAsFile(tempFilePath);

                if (_securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip || _securityForReceivedMail.IsWarnLinkFileInTheZip || _securityForReceivedMail.IsWarnOneFileInTheZip || _securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip)
                {
                    var zipTools = new ZipFileHandler();
                    var izEncryptedZip = zipTools.CheckZipIsEncryptedAndGetIncludeExtensions(tempFilePath);

                    //暗号化ZIPファイルの場合の警告
                    if (_securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip && izEncryptedZip)
                    {
                        var dialogResult = MessageBox.Show(Properties.Resources.AttatchmentIsEncryptedZip + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                        if (dialogResult == MessageBoxResult.Yes)
                        {
                            //Open file.
                        }
                        else
                        {
                            cancel = true;
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch (Exception)
                            {
                                //Do Nothing.
                            }
                            return;
                        }
                    }

                    //Zip内にlinkファイルがある場合の警告
                    if (_securityForReceivedMail.IsWarnLinkFileInTheZip)
                    {
                        if (zipTools.IncludeExtensions.Contains(".link"))
                        {
                            var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachmentZip_link + Environment.NewLine + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (dialogResult == MessageBoxResult.Yes)
                            {
                                //Open file.
                            }
                            else
                            {
                                cancel = true;
                                try
                                {
                                    File.Delete(tempFilePath);
                                }
                                catch (Exception)
                                {
                                    //Do Nothing.
                                }
                                return;
                            }
                        }
                    }

                    //Zip内にOneNoteファイルがある場合の警告
                    if (_securityForReceivedMail.IsWarnOneFileInTheZip)
                    {
                        if (zipTools.IncludeExtensions.Contains(".one"))
                        {
                            var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachmentZip_one + Environment.NewLine + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (dialogResult == MessageBoxResult.Yes)
                            {
                                //Open file.
                            }
                            else
                            {
                                cancel = true;
                                try
                                {
                                    File.Delete(tempFilePath);
                                }
                                catch (Exception)
                                {
                                    //Do Nothing.
                                }
                                return;
                            }
                        }
                    }

                    //Zip内にマクロ付きOfficeファイルがある場合の警告
                    if (_securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip)
                    {
                        if (zipTools.IncludeExtensions.Contains(".docm") | zipTools.IncludeExtensions.Contains(".xlsm") | zipTools.IncludeExtensions.Contains(".pptm"))
                        {
                            var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachmentZip_macro + Environment.NewLine + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (dialogResult == MessageBoxResult.Yes)
                            {
                                //Open file.
                            }
                            else
                            {
                                cancel = true;
                                try
                                {
                                    File.Delete(tempFilePath);
                                }
                                catch (Exception)
                                {
                                    //Do Nothing.
                                }
                                return;
                            }
                        }
                    }
                }

                //Officeファイル内にマクロが含まれている場合の警告
                if (_securityForReceivedMail.IsWarnBeforeOpeningAttachmentsThatContainMacros)
                {
                    if (OfficeFileHandler.CheckOfficeFileHasVbProject(tempFilePath, Path.GetExtension(attachment.FileName).ToLower()))
                    {
                        var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachment_macro + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                        if (dialogResult == MessageBoxResult.Yes)
                        {
                            //Open file.
                        }
                        else
                        {
                            cancel = true;
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch (Exception)
                            {
                                //Do Nothing.
                            }
                            return;
                        }
                    }
                }

                if (true)
                {

                }
                try
                {
                    File.Delete(tempFilePath);
                }
                catch (Exception)
                {
                    //Do Nothing.
                }
            }
        }

        /// <summary>
        /// 送信トレイのメールアイテムを開く際の警告。
        /// </summary>
        /// <param name="inspector">Inspector</param>
        private void OpenOutboxItemInspector(Outlook.Inspector inspector)
        {
            if (!(inspector.CurrentItem is Outlook.MailItem currentItem)) return;

            //送信保留中のメールのみ警告対象とする。
            if (currentItem.Submitted)
            {
                _ = MessageBox.Show(Properties.Resources.CanceledSendingMailMessage, Properties.Resources.CanceledSendingMail, MessageBoxButton.OK, MessageBoxImage.Warning);

                //再編集のため、配信指定日時をクリアする。
                currentItem.DeferredDeliveryTime = new DateTime(4501, 1, 1, 0, 0, 0);
                currentItem.Save();
            }

            ((Outlook.InspectorEvents_Event)inspector).Close += () =>
            {
                if (currentItem != null)
                {
                    _ = Marshal.ReleaseComObject(currentItem);
                    currentItem = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }

                _ = Marshal.ReleaseComObject(inspector);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            };
        }

        /// <summary>
        /// メール送信時(送信ボタン押下時)に確認画面を生成する。
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="cancel">Cancel</param>
        private void Application_ItemSend(object item, ref bool cancel)
        {
            //Outlook起動後にユーザが設定を変更する可能性があるため、毎回ユーザ設定をロード。
            LoadGeneralSetting(false);
            if (!(_generalSetting.LanguageCode is null))
            {
                ResourceService.Instance.ChangeCulture(_generalSetting.LanguageCode);
            }

            var autoAddMessageSetting = new AutoAddMessage();
            var autoAddMessageSettingList = CsvFileHandler.ReadCsv<AutoAddMessage>(typeof(AutoAddMessageMap), "AutoAddMessage.csv");
            if (autoAddMessageSettingList.Count > 0) autoAddMessageSetting = autoAddMessageSettingList[0];

            //Moderationでの返信には何もしない。(キャンセルすると、承認や非承認ができなくなる場合があるため)
            if (((dynamic)item).MessageClass == "IPM.Note.Microsoft.Approval.Reply.Approve" || ((dynamic)item).MessageClass == "IPM.Note.Microsoft.Approval.Reply.Reject") return;

            var type = typeof(Outlook.MailItem);
            //何らかの問題で確認画面が表示されないと、意図せずメールが送られてしまう恐れがあるため、念のための処理。
            try
            {
                try
                {
                    //FIXME: 暫定処置。
                    //HACK: 添付ファイルをリンクとして添付する際に、メール本文が自動更新されない問題を回避するための処置。
                    //HACK: ※WordEditorで本文を編集すると、本文の更新処理が行われるため問題を回避できる。
                    //HACK: ※メールの文頭に半角スペースを挿入し、それを削除することで、本文の編集処理とさせる。
                    var mailItemWordEditor = (Word.Document)((dynamic)item).GetInspector.WordEditor;
                    var range = mailItemWordEditor.Range(0, 0);
                    range.InsertAfter(" ");
                    range = mailItemWordEditor.Range(0, 0);
                    _ = range.Delete();
                }
                catch (Exception)
                {
                    //Do nothing.
                }

                //「連絡先に登録された宛先はあらかじめチェックを自動付与する。」など連絡先が必要な機能が有効な場合、連絡先をまとめて取得する。
                Outlook.MAPIFolder contacts = null;
                if (_generalSetting.IsAutoCheckRegisteredInContacts || _generalSetting.IsWarningIfRecipientsIsNotRegistered || _generalSetting.IsProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    contacts = Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                }

                var generateCheckList = new GenerateCheckList();
                CheckList checklist;
                switch (item)
                {
                    //MailItem(通常のメール)とMeetingItem(会議招待)の場合にのみ動作させる。
                    case Outlook.MailItem mailItem:
                        type = typeof(Outlook.MailItem);
                        checklist = generateCheckList.GenerateCheckListFromMail(mailItem, _generalSetting, contacts, autoAddMessageSetting);
                        break;
                    case Outlook.MeetingItem meetingItem when _generalSetting.IsShowConfirmationAtSendMeetingRequest:
                        type = typeof(Outlook.MeetingItem);
                        checklist = generateCheckList.GenerateCheckListFromMail(meetingItem, _generalSetting, contacts, autoAddMessageSetting);
                        break;
                    case Outlook.MeetingItem _:
                        return;
                    case Outlook.TaskRequestItem taskRequestItem when _generalSetting.IsShowConfirmationAtSendTaskRequest:
                        type = typeof(Outlook.TaskRequestItem);
                        checklist = generateCheckList.GenerateCheckListFromMail(taskRequestItem, _generalSetting, contacts, autoAddMessageSetting);
                        break;
                    case Outlook.TaskRequestItem _:
                        return;
                    default:
                        return;
                }

                if (_generalSetting.IsAutoCheckIfAllRecipientsAreSameDomain)
                {
                    foreach (var to in checklist.ToAddresses.Where(to => !to.IsExternal))
                    {
                        to.IsChecked = true;
                    }

                    foreach (var cc in checklist.CcAddresses.Where(cc => !cc.IsExternal))
                    {
                        cc.IsChecked = true;
                    }

                    foreach (var bcc in checklist.BccAddresses.Where(bcc => !bcc.IsExternal))
                    {
                        bcc.IsChecked = true;
                    }
                }

                if (_generalSetting.IsEnableRecipientsAreSortedByDomain)
                {
                    checklist.ToAddresses = checklist.ToAddresses.OrderBy(x => x.MailAddress.Substring((int)Math.Sqrt(Math.Pow(x.MailAddress.IndexOf("@", StringComparison.Ordinal), 2)))).ToList();
                    checklist.CcAddresses = checklist.CcAddresses.OrderBy(x => x.MailAddress.Substring((int)Math.Sqrt(Math.Pow(x.MailAddress.IndexOf("@", StringComparison.Ordinal), 2)))).ToList();
                    checklist.BccAddresses = checklist.BccAddresses.OrderBy(x => x.MailAddress.Substring((int)Math.Sqrt(Math.Pow(x.MailAddress.IndexOf("@", StringComparison.Ordinal), 2)))).ToList();
                }

                if (checklist.IsCanNotSendMail)
                {
                    //送信禁止条件に該当するため、確認画面を表示するのではなく、送信禁止画面を表示する。
                    //このタイミングで落ちると、メールが送信されてしまうので、念のためのTry Catch。
                    try
                    {
                        _ = MessageBox.Show(checklist.CanNotSendMailMessage, Properties.Resources.SendingForbid, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (Exception)
                    {
                        //Do nothing.
                    }
                    finally
                    {
                        cancel = true;
                    }

                    cancel = true;
                }
                else if (IsShowConfirmationWindow(checklist))
                {
                    //OutlookのWindowを親として確認画面をモーダル表示。
                    var confirmationWindow = new ConfirmationWindow(checklist, item);
                    var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                    var outlookHandle = new NativeMethods(activeWindow).Handle;
                    _ = new WindowInteropHelper(confirmationWindow) { Owner = outlookHandle };

                    var dialogResult = confirmationWindow.ShowDialog() ?? false;

                    if (dialogResult)
                    {
                        //メール本文への文言の自動追加はメール送信時に実行する。
                        AutoAddMessageToBody(autoAddMessageSetting, item, type == typeof(Outlook.MailItem));

                        //Send Mail.
                    }
                    else
                    {
                        cancel = true;
                    }
                }
                else
                {
                    //メール本文への文言の自動追加はメール送信時に実行する。
                    AutoAddMessageToBody(autoAddMessageSetting, item, type == typeof(Outlook.MailItem));

                    //Send Mail.
                }
            }
            catch (Exception e)
            {
                var dialogResult = MessageBox.Show(Properties.Resources.IsCanNotShowConfirmation + Environment.NewLine + Properties.Resources.SendMailConfirmation + Environment.NewLine + Environment.NewLine + e.Message, Properties.Resources.IsCanNotShowConfirmation, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dialogResult == MessageBoxResult.Yes)
                {
                    //メール本文への文言の自動追加はメール送信時に実行する。
                    AutoAddMessageToBody(autoAddMessageSetting, item, type == typeof(Outlook.MailItem));

                    //Send Mail.
                }
                else
                {
                    cancel = true;
                }
            }
        }

        /// <summary>
        /// 受信メールの関するセキュリティ機能の設定を読み込む
        /// </summary>
        private void LoadSecurityForReceivedMail()
        {
            var securityForReceivedMail = CsvFileHandler.ReadCsv<SecurityForReceivedMail>(typeof(SecurityForReceivedMailMap), "SecurityForReceivedMail.csv").ToList();
            if (securityForReceivedMail.Count == 0) return;

            _securityForReceivedMail.IsEnableSecurityForReceivedMail = securityForReceivedMail[0].IsEnableSecurityForReceivedMail;
            _securityForReceivedMail.IsEnableAlertKeywordOfSubjectWhenOpeningMailsData = securityForReceivedMail[0].IsEnableAlertKeywordOfSubjectWhenOpeningMailsData;
            _securityForReceivedMail.IsEnableMailHeaderAnalysis = securityForReceivedMail[0].IsEnableMailHeaderAnalysis;
            _securityForReceivedMail.IsShowWarningWhenSpfFails = securityForReceivedMail[0].IsShowWarningWhenSpfFails;
            _securityForReceivedMail.IsShowWarningWhenDkimFails = securityForReceivedMail[0].IsShowWarningWhenDkimFails;
            _securityForReceivedMail.IsEnableWarningFeatureWhenOpeningAttachments = securityForReceivedMail[0].IsEnableWarningFeatureWhenOpeningAttachments;
            _securityForReceivedMail.IsWarnBeforeOpeningAttachments = securityForReceivedMail[0].IsWarnBeforeOpeningAttachments;
            _securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip = securityForReceivedMail[0].IsWarnBeforeOpeningEncryptedZip;
            _securityForReceivedMail.IsWarnLinkFileInTheZip = securityForReceivedMail[0].IsWarnLinkFileInTheZip;
            _securityForReceivedMail.IsWarnOneFileInTheZip = securityForReceivedMail[0].IsWarnOneFileInTheZip;
            _securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip = securityForReceivedMail[0].IsWarnOfficeFileWithMacroInTheZip;
            _securityForReceivedMail.IsWarnBeforeOpeningAttachmentsThatContainMacros = securityForReceivedMail[0].IsWarnBeforeOpeningAttachmentsThatContainMacros;
        }

        /// <summary>
        /// 受信したメールの件名の警告対象となる設定を読み込む。
        /// </summary>
        private void LoadAlertKeywordOfSubjectWhenOpeningMailsData()
        {
            var alertKeywordOfSubjectWhenOpeningMails = CsvFileHandler.ReadCsv<AlertKeywordOfSubjectWhenOpeningMail>(typeof(AlertKeywordOfSubjectWhenOpeningMailMap), "AlertKeywordOfSubjectWhenOpeningMailList.csv").Where(x => !string.IsNullOrEmpty(x.AlertKeyword));
            _alertKeywordOfSubjectWhenOpeningMail.AddRange(alertKeywordOfSubjectWhenOpeningMails);
        }

        /// <summary>
        /// 一般設定を設定ファイルから読み込む。
        /// </summary>
        /// <param name="isLaunch">Outlookの起動時か否か</param>
        private void LoadGeneralSetting(bool isLaunch)
        {
            var generalSetting = CsvFileHandler.ReadCsv<GeneralSetting>(typeof(GeneralSettingMap), "GeneralSetting.csv").ToList();
            if (generalSetting.Count == 0) return;

            _generalSetting.LanguageCode = generalSetting[0].LanguageCode;

            if (isLaunch) return;

            _generalSetting.EnableForgottenToAttachAlert = generalSetting[0].EnableForgottenToAttachAlert;
            _generalSetting.IsDoNotConfirmationIfAllRecipientsAreSameDomain = generalSetting[0].IsDoNotConfirmationIfAllRecipientsAreSameDomain;
            _generalSetting.IsDoDoNotConfirmationIfAllWhite = generalSetting[0].IsDoDoNotConfirmationIfAllWhite;
            _generalSetting.IsAutoCheckIfAllRecipientsAreSameDomain = generalSetting[0].IsAutoCheckIfAllRecipientsAreSameDomain;
            _generalSetting.IsShowConfirmationToMultipleDomain = generalSetting[0].IsShowConfirmationToMultipleDomain;
            _generalSetting.EnableGetContactGroupMembers = generalSetting[0].EnableGetContactGroupMembers;
            _generalSetting.EnableGetExchangeDistributionListMembers = generalSetting[0].EnableGetExchangeDistributionListMembers;
            _generalSetting.ContactGroupMembersAreWhite = generalSetting[0].ContactGroupMembersAreWhite;
            _generalSetting.ExchangeDistributionListMembersAreWhite = generalSetting[0].ExchangeDistributionListMembersAreWhite;
            _generalSetting.IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = generalSetting[0].IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles;
            _generalSetting.IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain = generalSetting[0].IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain;
            _generalSetting.IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain = generalSetting[0].IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain;
            _generalSetting.IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain = generalSetting[0].IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain;
            _generalSetting.IsEnableRecipientsAreSortedByDomain = generalSetting[0].IsEnableRecipientsAreSortedByDomain;
            _generalSetting.IsAutoAddSenderToBcc = generalSetting[0].IsAutoAddSenderToBcc;
            _generalSetting.IsAutoCheckRegisteredInContacts = generalSetting[0].IsAutoCheckRegisteredInContacts;
            _generalSetting.IsAutoCheckRegisteredInContactsAndMemberOfContactLists = generalSetting[0].IsAutoCheckRegisteredInContactsAndMemberOfContactLists;
            _generalSetting.IsCheckNameAndDomainsFromRecipients = generalSetting[0].IsCheckNameAndDomainsFromRecipients;
            _generalSetting.IsWarningIfRecipientsIsNotRegistered = generalSetting[0].IsWarningIfRecipientsIsNotRegistered;
            _generalSetting.IsProhibitsSendingMailIfRecipientsIsNotRegistered = generalSetting[0].IsProhibitsSendingMailIfRecipientsIsNotRegistered;
            _generalSetting.IsShowConfirmationAtSendMeetingRequest = generalSetting[0].IsShowConfirmationAtSendMeetingRequest;
            _generalSetting.IsAutoAddSenderToCc = generalSetting[0].IsAutoAddSenderToCc;
            _generalSetting.IsCheckNameAndDomainsIncludeSubject = generalSetting[0].IsCheckNameAndDomainsIncludeSubject;
            _generalSetting.IsCheckNameAndDomainsFromSubject = generalSetting[0].IsCheckNameAndDomainsFromSubject;
            _generalSetting.IsShowConfirmationAtSendTaskRequest = generalSetting[0].IsShowConfirmationAtSendTaskRequest;
            _generalSetting.IsAutoCheckAttachments = generalSetting[0].IsAutoCheckAttachments;
            _generalSetting.IsCheckKeywordAndRecipientsIncludeSubject = generalSetting[0].IsCheckKeywordAndRecipientsIncludeSubject;
        }

        /// <summary>
        /// 全てのチェック対象がチェックされているか否かの判定。(ホワイトリスト登録の宛先など、事前にチェックされている場合がある)
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <returns>全てのチェック対象がチェックされているか否か</returns>
        private bool IsAllChecked(CheckList checkList)
        {
            var isToAddressesCompleteChecked = checkList.ToAddresses.Count(x => x.IsChecked) == checkList.ToAddresses.Count;
            var isCcAddressesCompleteChecked = checkList.CcAddresses.Count(x => x.IsChecked) == checkList.CcAddresses.Count;
            var isBccAddressesCompleteChecked = checkList.BccAddresses.Count(x => x.IsChecked) == checkList.BccAddresses.Count;
            var isAlertsCompleteChecked = checkList.Alerts.Count(x => x.IsChecked) == checkList.Alerts.Count;
            var isAttachmentsCompleteChecked = checkList.Attachments.Count(x => x.IsChecked) == checkList.Attachments.Count;

            return isToAddressesCompleteChecked && isCcAddressesCompleteChecked && isBccAddressesCompleteChecked && isAlertsCompleteChecked && isAttachmentsCompleteChecked;
        }

        /// <summary>
        /// 全ての宛先が内部(社内)ドメインであるか否かの判定。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <returns>全ての宛先が内部(社内)ドメインであるか否か</returns>
        private bool IsAllRecipientsAreSameDomain(CheckList checkList)
        {
            var isAllToRecipientsAreSameDomain = checkList.ToAddresses.Count(x => !x.IsExternal) == checkList.ToAddresses.Count;
            var isAllCcRecipientsAreSameDomain = checkList.CcAddresses.Count(x => !x.IsExternal) == checkList.CcAddresses.Count;
            var isAllBccRecipientsAreSameDomain = checkList.BccAddresses.Count(x => !x.IsExternal) == checkList.BccAddresses.Count;

            return isAllToRecipientsAreSameDomain && isAllCcRecipientsAreSameDomain && isAllBccRecipientsAreSameDomain;
        }

        /// <summary>
        /// 送信前の確認画面の表示有無を判定。
        /// </summary>
        /// <param name="checklist">CheckList</param>
        /// <returns>送信前の確認画面の表示有無</returns>
        private bool IsShowConfirmationWindow(CheckList checklist)
        {
            if (checklist.RecipientExternalDomainNumAll >= 2 && _generalSetting.IsShowConfirmationToMultipleDomain)
            {
                //全ての宛先が確認対象だが、複数のドメインが宛先に含まれる場合は確認画面を表示するオプションが有効かつその状態のため、スキップしない。
                //他の判定より優先されるため、常に先に確認して、先にreturnする。
                return true;
            }

            if (_generalSetting.IsDoNotConfirmationIfAllRecipientsAreSameDomain && IsAllRecipientsAreSameDomain(checklist))
            {
                //全ての受信者が送信者と同一ドメインの場合に確認画面を表示しないオプションが有効かつその状態のためスキップ。
                return false;
            }

            if (checklist.ToAddresses.Count(x => x.IsSkip) == checklist.ToAddresses.Count && checklist.CcAddresses.Count(x => x.IsSkip) == checklist.CcAddresses.Count && checklist.BccAddresses.Count(x => x.IsSkip) == checklist.BccAddresses.Count)
            {
                //全ての宛先が確認画面スキップ対象のためスキップ。
                return false;
            }

            if (_generalSetting.IsDoDoNotConfirmationIfAllWhite && IsAllChecked(checklist))
            {
                //全てにチェックが入った状態の場合に確認画面を表示しないオプションが有効かつその状態のためスキップ。
                return false;
            }

            //どのようなオプション条件にも該当しないため、通常通り確認画面を表示する。
            return true;
        }

        /// <summary>
        /// メール本文への文言の自動追加
        /// </summary>
        /// <param name="autoAddMessageSetting"></param>
        /// <param name="item"></param>
        /// <param name="isMailItem"></param>
        private void AutoAddMessageToBody(AutoAddMessage autoAddMessageSetting, object item, bool isMailItem)
        {
            //一旦、通常のメールのみ対象とする。
            if (!isMailItem) return;

            if (autoAddMessageSetting.IsAddToStart)
            {
                var mailItemWordEditor = (Word.Document)((dynamic)item).GetInspector.WordEditor;
                var range = mailItemWordEditor.Range(0, 0);
                range.InsertBefore(autoAddMessageSetting.MessageOfAddToStart + Environment.NewLine + Environment.NewLine);
            }

            if (autoAddMessageSetting.IsAddToEnd)
            {
                var mailItemWordEditor = (Word.Document)((dynamic)item).GetInspector.WordEditor;
                var range = mailItemWordEditor.Range();
                range.InsertAfter(Environment.NewLine + Environment.NewLine + autoAddMessageSetting.MessageOfAddToEnd);
            }
        }

        #region VSTO generated code

        private void InternalStartup() => Startup += ThisAddIn_Startup;

        #endregion
    }
}