using OutlookOkan.CsvTools;
using OutlookOkan.Models;
using OutlookOkan.Services;
using OutlookOkan.Types;
using OutlookOkan.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using MessageBox = System.Windows.MessageBox;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan
{
    public partial class ThisAddIn
    {
        private string _language = "NotSet";
        private bool _isDoNotConfirmationIfAllRecipientsAreSameDomain;
        private bool _isDoDoNotConfirmationIfAllWhite;
        private bool _isAutoCheckIfAllRecipientsAreSameDomain;
        private bool _isShowConfirmationToMultipleDomain;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //ユーザ設定をロード(このタイミングでロードしておかないと、リボンの言語を変更できない。
            LoadSetting();
            if (_language != "NotSet")
            {
                ResourceService.Instance.ChangeCulture(_language);
            }

            Application.ItemSend += Application_ItemSend;
        }

        //TODO 分かりやすくするために機能を分離する。
        private void Application_ItemSend(object item, ref bool cancel)
        {
            //何らかの問題で確認画面が表示されないと、意図せずメールが送られてしまう恐れがあるため、念のための処理を入れておく。
            try
            {
                var generateCheckList = new GenerateCheckList();
                var checklist = generateCheckList.GenerateCheckListFromMail(item as Outlook._MailItem);

                //Outlook起動後にユーザが設定を変更する可能性があるため、毎回ユーザ設定をロード
                LoadSetting();
                if (_language != "NotSet")
                {
                    ResourceService.Instance.ChangeCulture(_language);
                }

                //送信先と同一のドメインはあらかじめチェックを入れるオプションが有効な場合、それをする。
                if (_isAutoCheckIfAllRecipientsAreSameDomain)
                {
                    foreach (var to in checklist.ToAddresses)
                    {
                        if (!to.IsExternal)
                        {
                            to.IsChecked = true;
                        }
                    }

                    foreach (var cc in checklist.CcAddresses)
                    {
                        if (!cc.IsExternal)
                        {
                            cc.IsChecked = true;
                        }
                    }

                    foreach (var bcc in checklist.BccAddresses)
                    {
                        if (!bcc.IsExternal)
                        {
                            bcc.IsChecked = true;
                        }
                    }
                }

                //送信禁止フラグの確認
                if (checklist.IsCanNotSendMail)
                {
                    //送信禁止条件に該当するため、確認画面を表示するのではなく、送信禁止画面を表示する。
                    //このタイミングで落ちると、メールが送信されてしまうので、念のためのTry Catch.
                    try
                    {
                        MessageBox.Show(checklist.CanNotSendMailMessage, Properties.Resources.SendingForbid,
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (Exception)
                    {
                        //Do Noting.
                    }
                    finally
                    {
                        cancel = true;
                    }

                    cancel = true;
                }
                //確認画面の表示条件に合致していたら
                else if(IsShowConfirmationWindow(checklist))
                {
                    var confirmationWindow = new ConfirmationWindow(checklist);
                    var dialogResult = confirmationWindow.ShowDialog();

                    if (dialogResult == true)
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
            catch (Exception)
            {
                var dialogResult = MessageBox.Show(Properties.Resources.SendMailConfirmation, Properties.Resources.IsCanNotShowConfirmation, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dialogResult == MessageBoxResult.Yes)
                {
                    //SendMail
                }
                else
                {
                    cancel = true;
                }
            }
        }

        private void LoadSetting()
        {
            var generalSetting = new List<GeneralSetting>();
            var readCsv = new ReadAndWriteCsv("GeneralSetting.csv");
            foreach (var data in readCsv.GetCsvRecords<GeneralSetting>(readCsv.LoadCsv<GeneralSettingMap>()))
            {
                generalSetting.Add((data));
            }

            if (generalSetting.Count != 0)
            {
                _language = generalSetting[0].LanguageCode;
                _isDoNotConfirmationIfAllRecipientsAreSameDomain = generalSetting[0].IsDoNotConfirmationIfAllRecipientsAreSameDomain;
                _isDoDoNotConfirmationIfAllWhite = generalSetting[0].IsDoDoNotConfirmationIfAllWhite;
                _isAutoCheckIfAllRecipientsAreSameDomain = generalSetting[0].IsAutoCheckIfAllRecipientsAreSameDomain;
                _isShowConfirmationToMultipleDomain = generalSetting[0].IsShowConfirmationToMultipleDomain;
            }
        }

        private bool IsAllChedked(CheckList checkLlist)
        {
            var isToAddressesCompleteChecked = checkLlist.ToAddresses.Count(x => x.IsChecked) == checkLlist.ToAddresses.Count;
            var isCcAddressesCompleteChecked = checkLlist.CcAddresses.Count(x => x.IsChecked) == checkLlist.CcAddresses.Count;
            var isBccAddressesCompleteChecked = checkLlist.BccAddresses.Count(x => x.IsChecked) == checkLlist.BccAddresses.Count;
            var isAlertsCompleteChecked = checkLlist.Alerts.Count(x => x.IsChecked) == checkLlist.Alerts.Count;
            var isAttachmentsCompleteChecked = checkLlist.Attachments.Count(x => x.IsChecked) == checkLlist.Attachments.Count;

            return isToAddressesCompleteChecked && isCcAddressesCompleteChecked && isBccAddressesCompleteChecked && isAlertsCompleteChecked && isAttachmentsCompleteChecked;
        }

        private bool IsAllRecipientsAreSameDomain(CheckList checkLlist)
        {
            var isAllToRecipientsAreSameDomain = checkLlist.ToAddresses.Count(x => !x.IsExternal) == checkLlist.ToAddresses.Count;
            var isAllCcRecipientsAreSameDomain = checkLlist.CcAddresses.Count(x => !x.IsExternal) == checkLlist.CcAddresses.Count;
            var isAllBccRecipientsAreSameDomain = checkLlist.BccAddresses.Count(x => !x.IsExternal) == checkLlist.BccAddresses.Count;

            return isAllToRecipientsAreSameDomain && isAllCcRecipientsAreSameDomain && isAllBccRecipientsAreSameDomain;
        }

        private bool IsShowConfirmationWindow(CheckList checklist)
        {
            if (checklist.RecipientExternalDomainNum >= 2 && _isShowConfirmationToMultipleDomain)
            {
                //全ての宛先が確認対象だが、複数のドメインが宛先に含まれる場合は確認画面を表示するオプションが有効かつその状態のため、スキップしない。
                //他の判定より優先されるために先人確認して、先にretrunする。
                return true;
            }

            if (_isDoNotConfirmationIfAllRecipientsAreSameDomain && IsAllRecipientsAreSameDomain(checklist))
            {
                //全ての受信者が送信者と同一ドメインの場合に確認画面を表示しないオプションが有効かつその状態のためスキップ.
                return false;
            }
            
            if (checklist.ToAddresses.Count(x => x.IsSkip) == checklist.ToAddresses.Count && checklist.CcAddresses.Count(x => x.IsSkip) == checklist.CcAddresses.Count && checklist.BccAddresses.Count(x => x.IsSkip) == checklist.BccAddresses.Count)
            {
                //全ての宛先が確認画面スキップ対象のためスキップ。
                return false;
            }

            if (_isDoDoNotConfirmationIfAllWhite && IsAllChedked(checklist))
            { 
                //全てにチェックが入った状態の場合に確認画面を表示しないオプションが有効かつその状態のためスキップ.
                return false;
            }

            //どのようなオプション条件にも該当しないため、通常通り確認画面を表示する。
            return true;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
        }

        #endregion
    }
}