using OutlookOkan.CsvTools;
using OutlookOkan.Helpers;
using OutlookOkan.Models;
using OutlookOkan.Services;
using OutlookOkan.Types;
using OutlookOkan.Views;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Interop;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan
{
    public partial class ThisAddIn
    {
        private readonly GeneralSetting _generalSetting = new GeneralSetting();
        private Outlook.Inspectors _inspectors;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        /// <summary>
        /// アドインのロード時(Outlookの起動時)の処理。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //ユーザ設定をロード。(このタイミングでロードしないと、リボンの表示言語を変更できない。)
            LoadGeneralSetting(true);
            if (!(_generalSetting.LanguageCode is null))
            {
                ResourceService.Instance.ChangeCulture(_generalSetting.LanguageCode);
            }

            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += OpenOutboxItemInspector;

            Application.ItemSend += Application_ItemSend;
        }

        //送信トレイのファイルを開く際の警告。
        private void OpenOutboxItemInspector(Outlook.Inspector inspector)
        {
            if (!(inspector.CurrentItem is Outlook._MailItem)) return;
            if (!(inspector.CurrentItem is Outlook.MailItem mailItem)) return;

            if (mailItem.Submitted)
            {
                MessageBox.Show(Properties.Resources.CanceledSendingMailMessage, Properties.Resources.CanceledSendingMail, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            //MailItemにキャストできないものは会議招待などメールではないものなので、何もしない。
            if (!(item is Outlook._MailItem)) return;

            //何らかの問題で確認画面が表示されないと、意図せずメールが送られてしまう恐れがあるため、念のための処理。
            try
            {
                //Outlook起動後にユーザが設定を変更する可能性があるため、毎回ユーザ設定をロード。
                LoadGeneralSetting(false);
                if (!(_generalSetting.LanguageCode is null))
                {
                    ResourceService.Instance.ChangeCulture(_generalSetting.LanguageCode);
                }

                var generateCheckList = new GenerateCheckList();
                var checklist = generateCheckList.GenerateCheckListFromMail((Outlook._MailItem)item, _generalSetting);

                //送信先と同一のドメインはあらかじめチェックを入れるオプションが有効な場合、それを実行。
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

                //送信禁止フラグの確認。
                if (checklist.IsCanNotSendMail)
                {
                    //送信禁止条件に該当するため、確認画面を表示するのではなく、送信禁止画面を表示する。
                    //このタイミングで落ちると、メールが送信されてしまうので、念のためのTry Catch。
                    try
                    {
                        MessageBox.Show(checklist.CanNotSendMailMessage, Properties.Resources.SendingForbid, MessageBoxButton.OK, MessageBoxImage.Error);
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
                //確認画面の表示条件に合致しいた場合。
                else if (IsShowConfirmationWindow(checklist))
                {
                    //OutlookのWindowを親として確認画面をモーダル表示。
                    var confirmationWindow = new ConfirmationWindow(checklist, (Outlook._MailItem)item);
                    var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                    var outlookHandle = new NativeMethods(activeWindow).Handle;
                    _ = new WindowInteropHelper(confirmationWindow) { Owner = outlookHandle };

                    var dialogResult = confirmationWindow.ShowDialog() ?? false;

                    if (dialogResult)
                    {
                        //Send Mail.
                    }
                    else
                    {
                        cancel = true;
                    }
                }
                else
                {
                    //Send Mail.
                }
            }
            catch (Exception e)
            {
                var dialogResult = MessageBox.Show(Properties.Resources.IsCanNotShowConfirmation + Environment.NewLine + Properties.Resources.SendMailConfirmation + Environment.NewLine + Environment.NewLine + e.Message, Properties.Resources.IsCanNotShowConfirmation, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dialogResult == MessageBoxResult.Yes)
                {
                    //Send Mail.
                }
                else
                {
                    cancel = true;
                }
            }
        }

        /// <summary>
        /// 一般設定を設定ファイルから読み込む。
        /// </summary>
        /// <param name="isLaunch">Outlookの起動時か否か。</param>
        private void LoadGeneralSetting(bool isLaunch)
        {
            var readCsv = new ReadAndWriteCsv("GeneralSetting.csv");
            var generalSetting = readCsv.GetCsvRecords<GeneralSetting>(readCsv.LoadCsv<GeneralSettingMap>()).ToList();

            if (generalSetting.Count == 0) return;

            _generalSetting.LanguageCode = generalSetting[0].LanguageCode;
            if (isLaunch) return;

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
        }

        /// <summary>
        /// 全てのチェック対象がチェックされているか否かの判定。(ホワイトリスト登録の宛先など、事前にチェックされている場合がある)
        /// </summary>
        /// <param name="checkList"></param>
        /// <returns>全てのチェック対象がチェックされているか否か。</returns>
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
        /// <returns>全ての宛先が内部(社内)ドメインであるか否か。</returns>
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
        /// <returns>送信前の確認画面の表示有無。</returns>
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

        #region VSTO generated code

        private void InternalStartup() => Startup += ThisAddIn_Startup;

        #endregion
    }
}