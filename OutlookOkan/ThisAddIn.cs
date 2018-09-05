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

        private void Application_ItemSend(object item, ref bool cancel)
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
                foreach(var to in checklist.ToAddresses)
                {
                    if (!to.IsExternal)
                    {
                        to.IsChecked = true;
                    }
                }

                foreach(var cc in checklist.CcAddresses)
                {
                    if (!cc.IsExternal)
                    {
                        cc.IsChecked = true;
                    }
                }

                foreach(var bcc in checklist.BccAddresses)
                {
                    if (!bcc.IsExternal)
                    {
                        bcc.IsChecked = true;
                    }
                }
            }

            if(_isDoNotConfirmationIfAllRecipientsAreSameDomain && IsAllRecipientsAreSameDomain(checklist))
            {
                //全ての受信者が送信者と同一ドメインの場合に確認画面を表示しないオプションが有効かつその状態のためreturn.
                return;
            }


            if (_isDoDoNotConfirmationIfAllWhite && IsAllChedked(checklist))
            {
                //全てにチェックが入った状態の場合に確認画面を表示しないオプションが有効かつその状態のためreturn.
                return;
            }

            //送信禁止フラグの確認
            if (checklist.IsCanNotSendMail)
            {
                //このタイミングで落ちると、メールが送信されてしまうので、念のため。
                try
                {
                    MessageBox.Show(checklist.CanNotSendMailMessage, Properties.Resources.SendingForbid, MessageBoxButton.OK, MessageBoxImage.Error);
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
            else
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