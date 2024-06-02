using OutlookOkan.Handlers;
using OutlookOkan.Services;
using OutlookOkan.Types;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using Languages = OutlookOkan.Types.Languages;
using Task = System.Threading.Tasks.Task;

namespace OutlookOkan.ViewModels
{
    internal sealed class SettingsWindowViewModel : ViewModelBase
    {
        internal SettingsWindowViewModel()
        {
            //Add button command.
            ImportWhiteList = new RelayCommand(ImportWhiteListFromCsv);
            ExportWhiteList = new RelayCommand(ExportWhiteListToCsv);

            ImportNameAndDomainsList = new RelayCommand(ImportNameAndDomainsFromCsv);
            ExportNameAndDomainsList = new RelayCommand(ExportNameAndDomainsToCsv);

            ImportKeywordAndRecipientsList = new RelayCommand(ImportKeywordAndRecipientsFromCsv);
            ExportKeywordAndRecipientsList = new RelayCommand(ExportKeywordAndRecipientsToCsv);

            ImportAlertKeywordAndMessagesList = new RelayCommand(ImportAlertKeywordAndMessagesFromCsv);
            ExportAlertKeywordAndMessagesList = new RelayCommand(ExportAlertKeywordAndMessagesToCsv);

            ImportAlertAddressesList = new RelayCommand(ImportAlertAddressesFromCsv);
            ExportAlertAddressesList = new RelayCommand(ExportAlertAddressesToCsv);

            ImportAutoCcBccKeywordsList = new RelayCommand(ImportAutoCcBccKeywordsFromCsv);
            ExportAutoCcBccKeywordsList = new RelayCommand(ExportAutoCcBccKeywordsToCsv);

            ImportAutoCcBccRecipientsList = new RelayCommand(ImportAutoCcBccRecipientsFromCsv);
            ExportAutoCcBccRecipientsList = new RelayCommand(ExportAutoCcBccRecipientsToCsv);

            ImportAutoCcBccAttachedFilesList = new RelayCommand(ImportAutoCcBccAttachedFilesFromCsv);
            ExportAutoCcBccAttachedFilesList = new RelayCommand(ExportAutoCcBccAttachedFilesToCsv);

            ImportDeferredDeliveryMinutesList = new RelayCommand(ImportDeferredDeliveryMinutesFromCsv);
            ExportDeferredDeliveryMinutesList = new RelayCommand(ExportDeferredDeliveryMinutesToCsv);

            ImportInternalDomainList = new RelayCommand(ImportInternalDomainListFromCsv);
            ExportInternalDomainList = new RelayCommand(ExportInternalDomainListToCsv);

            ImportAlertKeywordAndMessagesForSubjectList = new RelayCommand(ImportAlertKeywordAndMessagesForSubjectFromCsv);
            ExportAlertKeywordAndMessagesForSubjectList = new RelayCommand(ExportAlertKeywordAndMessagesForSubjectToCsv);

            ImportRecipientsAndAttachmentsName = new RelayCommand(ImportRecipientsAndAttachmentsNameFromCsv);
            ExportRecipientsAndAttachmentsName = new RelayCommand(ExportRecipientsAndAttachmentsNameToCsv);

            ImportAttachmentProhibitedRecipients = new RelayCommand(ImportAttachmentProhibitedRecipientsFromCsv);
            ExportAttachmentProhibitedRecipients = new RelayCommand(ExportAttachmentProhibitedRecipientsToCsv);

            ImportAttachmentAlertRecipients = new RelayCommand(ImportAttachmentAlertRecipientsFromCsv);
            ExportAttachmentAlertRecipients = new RelayCommand(ExportAttachmentAlertRecipientsToCsv);

            ImportAlertKeywordOfSubjectWhenOpeningMailsList = new RelayCommand(ImportAlertKeywordOfSubjectWhenOpeningMailsFromCsv);
            ExportAlertKeywordOfSubjectWhenOpeningMailsList = new RelayCommand(ExportAlertKeywordOfSubjectWhenOpeningMailsToCsv);

            ImportAutoDeleteRecipientsList = new RelayCommand(ImportAutoDeleteRecipientsFromCsv);
            ExportAutoDeleteRecipientsList = new RelayCommand(ExportAutoDeleteRecipientsToCsv);

            //Load language code and name.
            var languages = new Languages();
            Languages = languages.Language;

            //Load settings from csv.
            LoadGeneralSettingData();
            LoadWhitelistData();
            LoadNameAndDomainsData();
            LoadKeywordAndRecipientsData();
            LoadAlertKeywordAndMessagesData();
            LoadAlertKeywordAndMessagesForSubjectData();
            LoadAlertAddressesData();
            LoadAutoCcBccKeywordsData();
            LoadAutoCcBccRecipientsData();
            LoadAutoCcBccAttachedFilesData();
            LoadDeferredDeliveryMinutesData();
            LoadInternalDomainListData();
            LoadExternalDomainsWarningAndAutoChangeToBccData();
            LoadAttachmentsSettingData();
            LoadRecipientsAndAttachmentsNameData();
            LoadAttachmentProhibitedRecipientsData();
            LoadAttachmentAlertRecipientsData();
            LoadForceAutoChangeRecipientsToBccData();
            LoadAlertKeywordOfSubjectWhenOpeningMailsData();
            LoadAutoDeleteRecipientsData();
            LoadAutoAddMessageData();
            LoadSecurityForReceivedMailData();
        }

        internal async Task SaveSettings()
        {
            IEnumerable<Task> saveTasks = new[]
                {
                    SaveGeneralSettingToCsv(),
                    SaveWhiteListToCsv(),
                    SaveNameAndDomainsToCsv(),
                    SaveKeywordAndRecipientsToCsv(),
                    SaveAlertKeywordAndMessageToCsv(),
                    SaveAlertKeywordAndMessageForSubjectToCsv(),
                    SaveAutoCcBccKeywordsToCsv(),
                    SaveAlertAddressesToCsv(),
                    SaveAutoCcBccRecipientsToCsv(),
                    SaveAutoCcBccAttachedFilesToCsv(),
                    SaveDeferredDeliveryMinutesToCsv(),
                    SaveInternalDomainListToCsv(),
                    SaveExternalDomainsWarningAndAutoChangeToBccToCsv(),
                    SaveAttachmentsSettingToCsv(),
                    SaveRecipientsAndAttachmentsNameToCsv(),
                    SaveAttachmentProhibitedRecipientsToCsv(),
                    SaveAttachmentAlertRecipientsToCsv(),
                    SaveForceAutoChangeRecipientsToBccToCsv(),
                    SaveAlertKeywordOfSubjectWhenOpeningMailToCsv(),
                    SaveAutoDeleteRecipientToCsv(),
                    SaveAutoAddMessageToCsv(),
                    SecurityForReceivedMailToCsv()
                };

            await Task.WhenAll(saveTasks);
        }

        #region Whitelist

        public ICommand ImportWhiteList { get; }
        public ICommand ExportWhiteList { get; }

        private void LoadWhitelistData()
        {
            var whiteList = CsvFileHandler.ReadCsv<Whitelist>(typeof(WhitelistMap), "Whitelist.csv");
            foreach (var data in whiteList.Where(x => !string.IsNullOrEmpty(x.WhiteName)))
            {
                Whitelist.Add(data);
            }
        }

        private async Task SaveWhiteListToCsv()
        {
            var list = Whitelist.Where(x => !string.IsNullOrEmpty(x.WhiteName)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(WhitelistMap), "Whitelist.csv", list));
        }

        private void ImportWhiteListFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<Whitelist>(typeof(WhitelistMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.WhiteName)))
                {
                    Whitelist.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportWhiteListToCsv()
        {
            var list = Whitelist.Where(x => !string.IsNullOrEmpty(x.WhiteName)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(WhitelistMap), list, "Whitelist.csv");
        }

        private ObservableCollection<Whitelist> _whitelist = new ObservableCollection<Whitelist>();
        public ObservableCollection<Whitelist> Whitelist
        {
            get => _whitelist;
            set
            {
                _whitelist = value;
                OnPropertyChanged(nameof(Whitelist));
            }
        }

        #endregion

        #region NameAndDomains

        public ICommand ImportNameAndDomainsList { get; }
        public ICommand ExportNameAndDomainsList { get; }

        private void LoadNameAndDomainsData()
        {
            var nameAndDomains = CsvFileHandler.ReadCsv<NameAndDomains>(typeof(NameAndDomainsMap), "NameAndDomains.csv");
            foreach (var data in nameAndDomains.Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)))
            {
                NameAndDomains.Add(data);
            }
        }

        private async Task SaveNameAndDomainsToCsv()
        {
            var list = NameAndDomains.Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(NameAndDomainsMap), "NameAndDomains.csv", list));
        }

        private void ImportNameAndDomainsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<NameAndDomains>(typeof(NameAndDomainsMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)))
                {
                    NameAndDomains.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportNameAndDomainsToCsv()
        {
            var list = NameAndDomains.Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(NameAndDomainsMap), list, "NameAndDomains.csv");
        }

        private ObservableCollection<NameAndDomains> _nameAndDomains = new ObservableCollection<NameAndDomains>();
        public ObservableCollection<NameAndDomains> NameAndDomains
        {
            get => _nameAndDomains;
            set
            {
                _nameAndDomains = value;
                OnPropertyChanged(nameof(NameAndDomains));
            }
        }

        #endregion

        #region KeywordAndRecipients

        public ICommand ImportKeywordAndRecipientsList { get; }
        public ICommand ExportKeywordAndRecipientsList { get; }

        private void LoadKeywordAndRecipientsData()
        {
            var keywordAndRecipients = CsvFileHandler.ReadCsv<KeywordAndRecipients>(typeof(KeywordAndRecipientsMap), "KeywordAndRecipientsList.csv");
            foreach (var data in keywordAndRecipients.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.Recipient)))
            {
                KeywordAndRecipients.Add(data);
            }
        }

        private async Task SaveKeywordAndRecipientsToCsv()
        {
            var list = KeywordAndRecipients.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(KeywordAndRecipientsMap), "KeywordAndRecipientsList.csv", list));
        }

        private void ImportKeywordAndRecipientsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<KeywordAndRecipients>(typeof(KeywordAndRecipientsMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.Recipient)))
                {
                    KeywordAndRecipients.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportKeywordAndRecipientsToCsv()
        {
            var list = KeywordAndRecipients.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(KeywordAndRecipientsMap), list, "KeywordAndRecipientsList.csv");
        }

        private ObservableCollection<KeywordAndRecipients> _keywordAndRecipients = new ObservableCollection<KeywordAndRecipients>();
        public ObservableCollection<KeywordAndRecipients> KeywordAndRecipients
        {
            get => _keywordAndRecipients;
            set
            {
                _keywordAndRecipients = value;
                OnPropertyChanged(nameof(KeywordAndRecipients));
            }
        }

        #endregion

        #region AlertKeywordAndMessagesForSubject

        public ICommand ImportAlertKeywordAndMessagesForSubjectList { get; }
        public ICommand ExportAlertKeywordAndMessagesForSubjectList { get; }

        private void LoadAlertKeywordAndMessagesForSubjectData()
        {
            var alertKeywordAndMessagesForSubject = CsvFileHandler.ReadCsv<AlertKeywordAndMessageForSubject>(typeof(AlertKeywordAndMessageForSubjectMap), "AlertKeywordAndMessageListForSubject.csv");
            foreach (var data in alertKeywordAndMessagesForSubject.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
            {
                AlertKeywordAndMessagesForSubject.Add(data);
            }
        }

        private async Task SaveAlertKeywordAndMessageForSubjectToCsv()
        {
            var list = AlertKeywordAndMessagesForSubject.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AlertKeywordAndMessageForSubjectMap), "AlertKeywordAndMessageListForSubject.csv", list));
        }

        private void ImportAlertKeywordAndMessagesForSubjectFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AlertKeywordAndMessageForSubject>(typeof(AlertKeywordAndMessageForSubjectMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
                {
                    AlertKeywordAndMessagesForSubject.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAlertKeywordAndMessagesForSubjectToCsv()
        {
            var list = AlertKeywordAndMessagesForSubject.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AlertKeywordAndMessageForSubjectMap), list, "AlertKeywordAndMessageListForSubject.csv");
        }

        private ObservableCollection<AlertKeywordAndMessageForSubject> _alertKeywordAndMessagesForSubject = new ObservableCollection<AlertKeywordAndMessageForSubject>();
        public ObservableCollection<AlertKeywordAndMessageForSubject> AlertKeywordAndMessagesForSubject
        {
            get => _alertKeywordAndMessagesForSubject;
            set
            {
                _alertKeywordAndMessagesForSubject = value;
                OnPropertyChanged(nameof(AlertKeywordAndMessagesForSubject));
            }
        }

        #endregion

        #region AlertKeywordAndMessages

        public ICommand ImportAlertKeywordAndMessagesList { get; }
        public ICommand ExportAlertKeywordAndMessagesList { get; }

        private void LoadAlertKeywordAndMessagesData()
        {
            var alertKeywordAndMessages = CsvFileHandler.ReadCsv<AlertKeywordAndMessage>(typeof(AlertKeywordAndMessageMap), "AlertKeywordAndMessageList.csv");
            foreach (var data in alertKeywordAndMessages.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
            {
                AlertKeywordAndMessages.Add(data);
            }
        }

        private async Task SaveAlertKeywordAndMessageToCsv()
        {
            var list = AlertKeywordAndMessages.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AlertKeywordAndMessageMap), "AlertKeywordAndMessageList.csv", list));
        }

        private void ImportAlertKeywordAndMessagesFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AlertKeywordAndMessage>(typeof(AlertKeywordAndMessageMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
                {
                    AlertKeywordAndMessages.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAlertKeywordAndMessagesToCsv()
        {
            var list = AlertKeywordAndMessages.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AlertKeywordAndMessageMap), list, "AlertKeywordAndMessageList.csv");
        }

        private ObservableCollection<AlertKeywordAndMessage> _alertKeywordAndMessages = new ObservableCollection<AlertKeywordAndMessage>();
        public ObservableCollection<AlertKeywordAndMessage> AlertKeywordAndMessages
        {
            get => _alertKeywordAndMessages;
            set
            {
                _alertKeywordAndMessages = value;
                OnPropertyChanged(nameof(AlertKeywordAndMessages));
            }
        }

        #endregion

        #region AlertAddresses

        public ICommand ImportAlertAddressesList { get; }
        public ICommand ExportAlertAddressesList { get; }

        private void LoadAlertAddressesData()
        {
            var alertAddresses = CsvFileHandler.ReadCsv<AlertAddress>(typeof(AlertAddressMap), "AlertAddressList.csv");
            foreach (var data in alertAddresses.Where(x => !string.IsNullOrEmpty(x.TargetAddress)))
            {
                AlertAddresses.Add(data);
            }
        }

        private async Task SaveAlertAddressesToCsv()
        {
            var list = AlertAddresses.Where(x => !string.IsNullOrEmpty(x.TargetAddress)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AlertAddressMap), "AlertAddressList.csv", list));
        }

        private void ImportAlertAddressesFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AlertAddress>(typeof(AlertAddressMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.TargetAddress)))
                {
                    AlertAddresses.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAlertAddressesToCsv()
        {
            var list = AlertAddresses.Where(x => !string.IsNullOrEmpty(x.TargetAddress)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AlertAddressMap), list, "AlertAddressList.csv");
        }

        private ObservableCollection<AlertAddress> _alertAddresses = new ObservableCollection<AlertAddress>();
        public ObservableCollection<AlertAddress> AlertAddresses
        {
            get => _alertAddresses;
            set
            {
                _alertAddresses = value;
                OnPropertyChanged(nameof(AlertAddresses));
            }
        }

        #endregion

        #region  AutoCcBccKeywords

        public ICommand ImportAutoCcBccKeywordsList { get; }
        public ICommand ExportAutoCcBccKeywordsList { get; }

        private void LoadAutoCcBccKeywordsData()
        {
            var autoCcBccKeywords = CsvFileHandler.ReadCsv<AutoCcBccKeyword>(typeof(AutoCcBccKeywordMap), "AutoCcBccKeywordList.csv");
            foreach (var data in autoCcBccKeywords.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.AutoAddAddress)))
            {
                AutoCcBccKeywords.Add(data);
            }
        }

        private async Task SaveAutoCcBccKeywordsToCsv()
        {
            var list = AutoCcBccKeywords.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AutoCcBccKeywordMap), "AutoCcBccKeywordList.csv", list));
        }

        private void ImportAutoCcBccKeywordsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AutoCcBccKeyword>(typeof(AutoCcBccKeywordMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.AutoAddAddress)))
                {
                    AutoCcBccKeywords.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAutoCcBccKeywordsToCsv()
        {
            var list = AutoCcBccKeywords.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AutoCcBccKeywordMap), list, "AutoCcBccKeywordList.csv");
        }

        private ObservableCollection<AutoCcBccKeyword> _autoCcBccKeywords = new ObservableCollection<AutoCcBccKeyword>();
        public ObservableCollection<AutoCcBccKeyword> AutoCcBccKeywords
        {
            get => _autoCcBccKeywords;
            set
            {
                _autoCcBccKeywords = value;
                OnPropertyChanged(nameof(AutoCcBccKeywords));
            }
        }

        #endregion

        #region AutoCcBccRecipient

        public ICommand ImportAutoCcBccRecipientsList { get; }
        public ICommand ExportAutoCcBccRecipientsList { get; }

        private void LoadAutoCcBccRecipientsData()
        {
            var autoCcBccRecipient = CsvFileHandler.ReadCsv<AutoCcBccRecipient>(typeof(AutoCcBccRecipientMap), "AutoCcBccRecipientList.csv");
            foreach (var data in autoCcBccRecipient.Where(x => !string.IsNullOrEmpty(x.TargetRecipient) && !string.IsNullOrEmpty(x.AutoAddAddress)))
            {
                AutoCcBccRecipients.Add(data);
            }
        }

        private async Task SaveAutoCcBccRecipientsToCsv()
        {
            var list = AutoCcBccRecipients.Where(x => !string.IsNullOrEmpty(x.TargetRecipient) && !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AutoCcBccRecipientMap), "AutoCcBccRecipientList.csv", list));
        }

        private void ImportAutoCcBccRecipientsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AutoCcBccRecipient>(typeof(AutoCcBccRecipientMap));

                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.TargetRecipient) && !string.IsNullOrEmpty(x.AutoAddAddress)))
                {
                    AutoCcBccRecipients.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAutoCcBccRecipientsToCsv()
        {
            var list = AutoCcBccRecipients.Where(x => !string.IsNullOrEmpty(x.TargetRecipient) && !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AutoCcBccRecipientMap), list, "AutoCcBccRecipientList.csv");
        }

        private ObservableCollection<AutoCcBccRecipient> _autoCcBccRecipients = new ObservableCollection<AutoCcBccRecipient>();
        public ObservableCollection<AutoCcBccRecipient> AutoCcBccRecipients
        {
            get => _autoCcBccRecipients;
            set
            {
                _autoCcBccRecipients = value;
                OnPropertyChanged(nameof(AutoCcBccRecipients));
            }
        }

        #endregion

        #region AutoCcBccAttachedFile

        public ICommand ImportAutoCcBccAttachedFilesList { get; }
        public ICommand ExportAutoCcBccAttachedFilesList { get; }

        private void LoadAutoCcBccAttachedFilesData()
        {
            var autoCcBccAttachedFile = CsvFileHandler.ReadCsv<AutoCcBccAttachedFile>(typeof(AutoCcBccAttachedFileMap), "AutoCcBccAttachedFileList.csv");
            foreach (var data in autoCcBccAttachedFile.Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)))
            {
                AutoCcBccAttachedFiles.Add(data);
            }
        }

        private async Task SaveAutoCcBccAttachedFilesToCsv()
        {
            var list = AutoCcBccAttachedFiles.Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AutoCcBccAttachedFileMap), "AutoCcBccAttachedFileList.csv", list));
        }

        private void ImportAutoCcBccAttachedFilesFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AutoCcBccAttachedFile>(typeof(AutoCcBccAttachedFileMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)))
                {
                    AutoCcBccAttachedFiles.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAutoCcBccAttachedFilesToCsv()
        {
            var list = AutoCcBccAttachedFiles.Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AutoCcBccAttachedFileMap), list, "AutoCcBccAttachedFileList.csv");
        }

        private ObservableCollection<AutoCcBccAttachedFile> _autoCcBccAttachedFiles = new ObservableCollection<AutoCcBccAttachedFile>();
        public ObservableCollection<AutoCcBccAttachedFile> AutoCcBccAttachedFiles
        {
            get => _autoCcBccAttachedFiles;
            set
            {
                _autoCcBccAttachedFiles = value;
                OnPropertyChanged(nameof(AutoCcBccAttachedFiles));
            }
        }

        #endregion

        #region DeferredDelivery

        public ICommand ImportDeferredDeliveryMinutesList { get; }
        public ICommand ExportDeferredDeliveryMinutesList { get; }

        private void LoadDeferredDeliveryMinutesData()
        {
            var deferredDeliveryMinutes = CsvFileHandler.ReadCsv<DeferredDeliveryMinutes>(typeof(DeferredDeliveryMinutesMap), "DeferredDeliveryMinutes.csv");
            foreach (var data in deferredDeliveryMinutes.Where(x => !string.IsNullOrEmpty(x.TargetAddress)))
            {
                DeferredDeliveryMinutes.Add(data);
            }
        }

        private async Task SaveDeferredDeliveryMinutesToCsv()
        {
            var list = DeferredDeliveryMinutes.Where(x => !string.IsNullOrEmpty(x.TargetAddress)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(DeferredDeliveryMinutesMap), "DeferredDeliveryMinutes.csv", list));
        }

        private void ImportDeferredDeliveryMinutesFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<DeferredDeliveryMinutes>(typeof(DeferredDeliveryMinutesMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.TargetAddress)))
                {
                    DeferredDeliveryMinutes.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportDeferredDeliveryMinutesToCsv()
        {
            var list = DeferredDeliveryMinutes.Where(x => !string.IsNullOrEmpty(x.TargetAddress)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(DeferredDeliveryMinutesMap), list, "DeferredDeliveryMinutes.csv");
        }

        private ObservableCollection<DeferredDeliveryMinutes> _deferredDeliveryMinutes = new ObservableCollection<DeferredDeliveryMinutes>();
        public ObservableCollection<DeferredDeliveryMinutes> DeferredDeliveryMinutes
        {
            get => _deferredDeliveryMinutes;
            set
            {
                _deferredDeliveryMinutes = value;
                OnPropertyChanged(nameof(DeferredDeliveryMinutes));
            }
        }

        #endregion

        #region InternalDomain

        public ICommand ImportInternalDomainList { get; }
        public ICommand ExportInternalDomainList { get; }

        private void LoadInternalDomainListData()
        {
            var internalDomainList = CsvFileHandler.ReadCsv<InternalDomain>(typeof(InternalDomainMap), "InternalDomainList.csv");
            foreach (var data in internalDomainList.Where(x => !string.IsNullOrEmpty(x.Domain)))
            {
                InternalDomainList.Add(data);
            }
        }

        private async Task SaveInternalDomainListToCsv()
        {
            var list = InternalDomainList.Where(x => !string.IsNullOrEmpty(x.Domain)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(InternalDomainMap), "InternalDomainList.csv", list));
        }

        private void ImportInternalDomainListFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<InternalDomain>(typeof(InternalDomainMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Domain)))
                {
                    InternalDomainList.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportInternalDomainListToCsv()
        {
            var list = InternalDomainList.Where(x => !string.IsNullOrEmpty(x.Domain)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(InternalDomainMap), list, "InternalDomainList.csv");
        }

        private ObservableCollection<InternalDomain> _internalDomainList = new ObservableCollection<InternalDomain>();
        public ObservableCollection<InternalDomain> InternalDomainList
        {
            get => _internalDomainList;
            set
            {
                _internalDomainList = value;
                OnPropertyChanged(nameof(InternalDomainList));
            }
        }

        #endregion

        #region ExternalDomainsWarningAndAutoChangeToBcc

        private void LoadExternalDomainsWarningAndAutoChangeToBccData()
        {
            var list = CsvFileHandler.ReadCsv<ExternalDomainsWarningAndAutoChangeToBcc>(typeof(ExternalDomainsWarningAndAutoChangeToBccMap), "ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
            if (list.Count == 0) return;

            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            _externalDomainsWarningAndAutoChangeToBcc.AddRange(list);

            //実際に使用するのは1行目の設定のみ
            TargetToAndCcExternalDomainsNum = _externalDomainsWarningAndAutoChangeToBcc[0].TargetToAndCcExternalDomainsNum;
            IsWarningWhenLargeNumberOfExternalDomains = _externalDomainsWarningAndAutoChangeToBcc[0].IsWarningWhenLargeNumberOfExternalDomains;
            IsProhibitedWhenLargeNumberOfExternalDomains = _externalDomainsWarningAndAutoChangeToBcc[0].IsProhibitedWhenLargeNumberOfExternalDomains;
            IsAutoChangeToBccWhenLargeNumberOfExternalDomains = _externalDomainsWarningAndAutoChangeToBcc[0].IsAutoChangeToBccWhenLargeNumberOfExternalDomains;
        }

        private async Task SaveExternalDomainsWarningAndAutoChangeToBccToCsv()
        {
            var tempExternalDomainsWarningAndAutoChangeToBcc = new List<ExternalDomainsWarningAndAutoChangeToBcc>
            {
                new ExternalDomainsWarningAndAutoChangeToBcc
                {
                    TargetToAndCcExternalDomainsNum = TargetToAndCcExternalDomainsNum,
                    IsWarningWhenLargeNumberOfExternalDomains = IsWarningWhenLargeNumberOfExternalDomains,
                    IsProhibitedWhenLargeNumberOfExternalDomains = IsProhibitedWhenLargeNumberOfExternalDomains,
                    IsAutoChangeToBccWhenLargeNumberOfExternalDomains = IsAutoChangeToBccWhenLargeNumberOfExternalDomains
                }
            };

            var list = tempExternalDomainsWarningAndAutoChangeToBcc.Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(ExternalDomainsWarningAndAutoChangeToBccMap), "ExternalDomainsWarningAndAutoChangeToBccSetting.csv", list));
        }

        private readonly List<ExternalDomainsWarningAndAutoChangeToBcc> _externalDomainsWarningAndAutoChangeToBcc = new List<ExternalDomainsWarningAndAutoChangeToBcc>();

        private int _targetToAndCcExternalDomainsNum = 10;
        public int TargetToAndCcExternalDomainsNum
        {
            get => _targetToAndCcExternalDomainsNum;
            set
            {
                _targetToAndCcExternalDomainsNum = value;
                OnPropertyChanged(nameof(TargetToAndCcExternalDomainsNum));
            }
        }

        private bool _isWarningWhenLargeNumberOfExternalDomains = true;
        public bool IsWarningWhenLargeNumberOfExternalDomains
        {
            get => _isWarningWhenLargeNumberOfExternalDomains;
            set
            {
                _isWarningWhenLargeNumberOfExternalDomains = value;
                OnPropertyChanged(nameof(IsWarningWhenLargeNumberOfExternalDomains));
            }
        }

        private bool _isProhibitedWhenLargeNumberOfExternalDomains;
        public bool IsProhibitedWhenLargeNumberOfExternalDomains
        {
            get => _isProhibitedWhenLargeNumberOfExternalDomains;
            set
            {
                _isProhibitedWhenLargeNumberOfExternalDomains = value;
                OnPropertyChanged(nameof(IsProhibitedWhenLargeNumberOfExternalDomains));
                OnPropertyChanged(nameof(IsWarningWhenLargeNumberOfExternalDomainsCheckBoxIsEnabled));
                OnPropertyChanged(nameof(IsAutoChangeToBccWhenLargeNumberOfExternalDomainsCheckBoxIsEnabled));
            }
        }

        private bool _isAutoChangeToBccWhenLargeNumberOfExternalDomains;
        public bool IsAutoChangeToBccWhenLargeNumberOfExternalDomains
        {
            get => _isAutoChangeToBccWhenLargeNumberOfExternalDomains;
            set
            {
                _isAutoChangeToBccWhenLargeNumberOfExternalDomains = value;
                OnPropertyChanged(nameof(IsAutoChangeToBccWhenLargeNumberOfExternalDomains));
            }
        }

        public bool IsWarningWhenLargeNumberOfExternalDomainsCheckBoxIsEnabled => !IsProhibitedWhenLargeNumberOfExternalDomains && !IsForceAutoChangeRecipientsToBcc;

        public bool IsAutoChangeToBccWhenLargeNumberOfExternalDomainsCheckBoxIsEnabled => !IsProhibitedWhenLargeNumberOfExternalDomains && !IsForceAutoChangeRecipientsToBcc;

        public bool TargetToAndCcExternalDomainsNumEnabledIsEnabled => !IsForceAutoChangeRecipientsToBcc;

        public bool IsProhibitedWhenLargeNumberOfExternalDomainsIsEnabled => !IsForceAutoChangeRecipientsToBcc;

        #endregion

        #region AttachmentsSetting

        private void LoadAttachmentsSettingData()
        {
            var list = CsvFileHandler.ReadCsv<AttachmentsSetting>(typeof(AttachmentsSettingMap), "AttachmentsSetting.csv");
            if (list.Count == 0) return;

            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            _attachmentsSetting.AddRange(list);

            //実際に使用するのは1行目の設定のみ
            IsWarningWhenEncryptedZipIsAttached = _attachmentsSetting[0].IsWarningWhenEncryptedZipIsAttached;
            IsProhibitedWhenEncryptedZipIsAttached = _attachmentsSetting[0].IsProhibitedWhenEncryptedZipIsAttached;
            IsEnableAllAttachedFilesAreDetectEncryptedZip = _attachmentsSetting[0].IsEnableAllAttachedFilesAreDetectEncryptedZip;
            IsAttachmentsProhibited = _attachmentsSetting[0].IsAttachmentsProhibited;
            IsWarningWhenAttachedRealFile = _attachmentsSetting[0].IsWarningWhenAttachedRealFile;
            IsEnableOpenAttachedFiles = _attachmentsSetting[0].IsEnableOpenAttachedFiles;
            TargetAttachmentFileExtensionOfOpen = _attachmentsSetting[0].TargetAttachmentFileExtensionOfOpen;
            IsMustOpenBeforeCheckTheAttachedFiles = _attachmentsSetting[0].IsMustOpenBeforeCheckTheAttachedFiles;
            IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain = _attachmentsSetting[0].IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain;

            if (string.IsNullOrEmpty(TargetAttachmentFileExtensionOfOpen)) TargetAttachmentFileExtensionOfOpen = ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx";
        }

        private async Task SaveAttachmentsSettingToCsv()
        {
            var tempAttachmentsSetting = new List<AttachmentsSetting>
            {
                new AttachmentsSetting
                {
                    IsWarningWhenEncryptedZipIsAttached = IsWarningWhenEncryptedZipIsAttached,
                    IsProhibitedWhenEncryptedZipIsAttached = IsProhibitedWhenEncryptedZipIsAttached,
                    IsEnableAllAttachedFilesAreDetectEncryptedZip = IsEnableAllAttachedFilesAreDetectEncryptedZip,
                    IsAttachmentsProhibited = IsAttachmentsProhibited,
                    IsWarningWhenAttachedRealFile = IsWarningWhenAttachedRealFile,
                    IsEnableOpenAttachedFiles = IsEnableOpenAttachedFiles,
                    TargetAttachmentFileExtensionOfOpen = TargetAttachmentFileExtensionOfOpen,
                    IsMustOpenBeforeCheckTheAttachedFiles = IsMustOpenBeforeCheckTheAttachedFiles,
                    IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain = IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain
                }
            };

            var list = tempAttachmentsSetting.Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AttachmentsSettingMap), "AttachmentsSetting.csv", list));
        }

        private readonly List<AttachmentsSetting> _attachmentsSetting = new List<AttachmentsSetting>();

        private bool _isWarningWhenEncryptedZipIsAttached;
        public bool IsWarningWhenEncryptedZipIsAttached
        {
            get => _isWarningWhenEncryptedZipIsAttached;
            set
            {
                _isWarningWhenEncryptedZipIsAttached = value;
                OnPropertyChanged(nameof(IsWarningWhenEncryptedZipIsAttached));
                OnPropertyChanged(nameof(IsEnableAllAttachedFilesAreDetectEncryptedZipCheckBoxIsEnabled));
            }
        }

        private bool _isProhibitedWhenEncryptedZipIsAttached;
        public bool IsProhibitedWhenEncryptedZipIsAttached
        {
            get => _isProhibitedWhenEncryptedZipIsAttached;
            set
            {
                _isProhibitedWhenEncryptedZipIsAttached = value;
                OnPropertyChanged(nameof(IsProhibitedWhenEncryptedZipIsAttached));
                OnPropertyChanged(nameof(IsEnableAllAttachedFilesAreDetectEncryptedZipCheckBoxIsEnabled));
                OnPropertyChanged(nameof(IsWarningWhenEncryptedZipIsAttachedCheckBoxIsEnabled));
            }
        }

        private bool _isEnableAllAttachedFilesAreDetectEncryptedZip;
        public bool IsEnableAllAttachedFilesAreDetectEncryptedZip
        {
            get => _isEnableAllAttachedFilesAreDetectEncryptedZip;
            set
            {
                _isEnableAllAttachedFilesAreDetectEncryptedZip = value;
                OnPropertyChanged(nameof(IsEnableAllAttachedFilesAreDetectEncryptedZip));
            }
        }

        private bool _isAttachmentsProhibited;
        public bool IsAttachmentsProhibited
        {
            get => _isAttachmentsProhibited;
            set
            {
                _isAttachmentsProhibited = value;
                OnPropertyChanged(nameof(IsAttachmentsProhibited));
                OnPropertyChanged(nameof(IsWarningWhenAttachedRealFileCheckBoxIsEnabled));
            }
        }

        private bool _isWarningWhenAttachedRealFile;
        public bool IsWarningWhenAttachedRealFile
        {
            get => _isWarningWhenAttachedRealFile;
            set
            {
                _isWarningWhenAttachedRealFile = value;
                OnPropertyChanged(nameof(IsWarningWhenAttachedRealFile));
            }
        }

        private bool _isEnableOpenAttachedFiles;
        public bool IsEnableOpenAttachedFiles
        {
            get => _isEnableOpenAttachedFiles;
            set
            {
                _isEnableOpenAttachedFiles = value;
                OnPropertyChanged(nameof(IsEnableOpenAttachedFiles));
            }
        }

        private string _targetAttachmentFileExtensionOfOpen = ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx";
        public string TargetAttachmentFileExtensionOfOpen
        {
            get => _targetAttachmentFileExtensionOfOpen;
            set
            {
                _targetAttachmentFileExtensionOfOpen = value;
                OnPropertyChanged(nameof(TargetAttachmentFileExtensionOfOpen));
            }
        }

        private bool _isMustOpenBeforeCheckTheAttachedFiles;
        public bool IsMustOpenBeforeCheckTheAttachedFiles
        {
            get => _isMustOpenBeforeCheckTheAttachedFiles;
            set
            {
                _isMustOpenBeforeCheckTheAttachedFiles = value;
                OnPropertyChanged(nameof(IsMustOpenBeforeCheckTheAttachedFiles));
                OnPropertyChanged(nameof(IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain));
            }
        }

        private bool _isIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain;
        public bool IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain
        {
            get => _isIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain;
            set
            {
                _isIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain = value;
                OnPropertyChanged(nameof(IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain));
            }
        }


        public bool IsWarningWhenEncryptedZipIsAttachedCheckBoxIsEnabled => !IsProhibitedWhenEncryptedZipIsAttached;

        public bool IsEnableAllAttachedFilesAreDetectEncryptedZipCheckBoxIsEnabled => IsWarningWhenEncryptedZipIsAttached || IsProhibitedWhenEncryptedZipIsAttached;

        public bool IsWarningWhenAttachedRealFileCheckBoxIsEnabled => !IsAttachmentsProhibited;

        #endregion

        #region RecipientsAndAttachmentsName

        public ICommand ImportRecipientsAndAttachmentsName { get; }
        public ICommand ExportRecipientsAndAttachmentsName { get; }

        private void LoadRecipientsAndAttachmentsNameData()
        {
            var recipientsAndAttachmentsName = CsvFileHandler.ReadCsv<RecipientsAndAttachmentsName>(typeof(RecipientsAndAttachmentsNameMap), "RecipientsAndAttachmentsName.csv");
            foreach (var data in recipientsAndAttachmentsName.Where(x => !string.IsNullOrEmpty(x.AttachmentsName) && !string.IsNullOrEmpty(x.Recipient)))
            {
                RecipientsAndAttachmentsName.Add(data);
            }
        }

        private async Task SaveRecipientsAndAttachmentsNameToCsv()
        {
            var list = RecipientsAndAttachmentsName.Where(x => !string.IsNullOrEmpty(x.AttachmentsName) && !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(RecipientsAndAttachmentsNameMap), "RecipientsAndAttachmentsName.csv", list));
        }

        private void ImportRecipientsAndAttachmentsNameFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<RecipientsAndAttachmentsName>(typeof(RecipientsAndAttachmentsNameMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.AttachmentsName) && !string.IsNullOrEmpty(x.Recipient)))
                {
                    RecipientsAndAttachmentsName.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportRecipientsAndAttachmentsNameToCsv()
        {
            var list = RecipientsAndAttachmentsName.Where(x => !string.IsNullOrEmpty(x.AttachmentsName) && !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(RecipientsAndAttachmentsNameMap), list, "RecipientsAndAttachmentsName.csv");
        }

        private ObservableCollection<RecipientsAndAttachmentsName> _recipientsAndAttachmentsName = new ObservableCollection<RecipientsAndAttachmentsName>();
        public ObservableCollection<RecipientsAndAttachmentsName> RecipientsAndAttachmentsName
        {
            get => _recipientsAndAttachmentsName;
            set
            {
                _recipientsAndAttachmentsName = value;
                OnPropertyChanged(nameof(RecipientsAndAttachmentsName));
            }
        }

        #endregion

        #region AttachmentProhibitedRecipients

        public ICommand ImportAttachmentProhibitedRecipients { get; }
        public ICommand ExportAttachmentProhibitedRecipients { get; }

        private void LoadAttachmentProhibitedRecipientsData()
        {
            var attachmentProhibitedRecipients = CsvFileHandler.ReadCsv<AttachmentProhibitedRecipients>(typeof(AttachmentProhibitedRecipientsMap), "AttachmentProhibitedRecipients.csv");
            foreach (var data in attachmentProhibitedRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)))
            {
                AttachmentProhibitedRecipients.Add(data);
            }
        }

        private async Task SaveAttachmentProhibitedRecipientsToCsv()
        {
            var list = AttachmentProhibitedRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AttachmentProhibitedRecipientsMap), "AttachmentProhibitedRecipients.csv", list));
        }

        private void ImportAttachmentProhibitedRecipientsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AttachmentProhibitedRecipients>(typeof(AttachmentProhibitedRecipientsMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Recipient)))
                {
                    AttachmentProhibitedRecipients.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAttachmentProhibitedRecipientsToCsv()
        {
            var list = AttachmentProhibitedRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AttachmentProhibitedRecipientsMap), list, "AttachmentProhibitedRecipients.csv");
        }

        private ObservableCollection<AttachmentProhibitedRecipients> _attachmentProhibitedRecipients = new ObservableCollection<AttachmentProhibitedRecipients>();
        public ObservableCollection<AttachmentProhibitedRecipients> AttachmentProhibitedRecipients
        {
            get => _attachmentProhibitedRecipients;
            set
            {
                _attachmentProhibitedRecipients = value;
                OnPropertyChanged(nameof(AttachmentProhibitedRecipients));
            }
        }

        #endregion

        #region AttachmentAlertRecipients

        public ICommand ImportAttachmentAlertRecipients { get; }
        public ICommand ExportAttachmentAlertRecipients { get; }

        private void LoadAttachmentAlertRecipientsData()
        {
            var attachmentAlertRecipients = CsvFileHandler.ReadCsv<AttachmentAlertRecipients>(typeof(AttachmentAlertRecipientsMap), "AttachmentAlertRecipients.csv");
            foreach (var data in attachmentAlertRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)))
            {
                AttachmentAlertRecipients.Add(data);
            }
        }

        private async Task SaveAttachmentAlertRecipientsToCsv()
        {
            var list = AttachmentAlertRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AttachmentAlertRecipientsMap), "AttachmentAlertRecipients.csv", list));
        }

        private void ImportAttachmentAlertRecipientsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AttachmentAlertRecipients>(typeof(AttachmentAlertRecipientsMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Recipient)))
                {
                    AttachmentAlertRecipients.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAttachmentAlertRecipientsToCsv()
        {
            var list = AttachmentAlertRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AttachmentAlertRecipientsMap), list, "AttachmentAlertRecipients.csv");
        }

        private ObservableCollection<AttachmentAlertRecipients> _attachmentAlertRecipients = new ObservableCollection<AttachmentAlertRecipients>();
        public ObservableCollection<AttachmentAlertRecipients> AttachmentAlertRecipients
        {
            get => _attachmentAlertRecipients;
            set
            {
                _attachmentAlertRecipients = value;
                OnPropertyChanged(nameof(AttachmentAlertRecipients));
            }
        }

        #endregion

        #region ForceAutoChangeRecipientsToBcc

        private void LoadForceAutoChangeRecipientsToBccData()
        {
            var list = CsvFileHandler.ReadCsv<ForceAutoChangeRecipientsToBcc>(typeof(ForceAutoChangeRecipientsToBccMap), "ForceAutoChangeRecipientsToBcc.csv");
            if (list.Count == 0) return;

            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            _forceAutoChangeRecipientsToBcc.AddRange(list);

            //実際に使用するのは1行目の設定のみ
            IsForceAutoChangeRecipientsToBcc = _forceAutoChangeRecipientsToBcc[0].IsForceAutoChangeRecipientsToBcc;
            ToRecipient = _forceAutoChangeRecipientsToBcc[0].ToRecipient;
            IsIncludeInternalDomain = _forceAutoChangeRecipientsToBcc[0].IsIncludeInternalDomain;
        }

        private async Task SaveForceAutoChangeRecipientsToBccToCsv()
        {
            var tempForceAutoChangeRecipientsToBcc = new List<ForceAutoChangeRecipientsToBcc>
            {
                new ForceAutoChangeRecipientsToBcc
                {
                    IsForceAutoChangeRecipientsToBcc = IsForceAutoChangeRecipientsToBcc,
                    ToRecipient = ToRecipient,
                    IsIncludeInternalDomain = IsIncludeInternalDomain
                }
            };

            var list = tempForceAutoChangeRecipientsToBcc.Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(ForceAutoChangeRecipientsToBccMap), "ForceAutoChangeRecipientsToBcc.csv", list));
        }

        private readonly List<ForceAutoChangeRecipientsToBcc> _forceAutoChangeRecipientsToBcc = new List<ForceAutoChangeRecipientsToBcc>();

        private bool _isForceAutoChangeRecipientsToBcc;
        public bool IsForceAutoChangeRecipientsToBcc
        {
            get => _isForceAutoChangeRecipientsToBcc;
            set
            {
                _isForceAutoChangeRecipientsToBcc = value;
                OnPropertyChanged(nameof(IsForceAutoChangeRecipientsToBcc));
                OnPropertyChanged(nameof(IsWarningWhenLargeNumberOfExternalDomainsCheckBoxIsEnabled));
                OnPropertyChanged(nameof(IsAutoChangeToBccWhenLargeNumberOfExternalDomainsCheckBoxIsEnabled));
                OnPropertyChanged(nameof(TargetToAndCcExternalDomainsNumEnabledIsEnabled));
                OnPropertyChanged(nameof(IsProhibitedWhenLargeNumberOfExternalDomainsIsEnabled));
            }
        }

        private string _toRecipient;
        public string ToRecipient
        {
            get => _toRecipient;
            set
            {
                _toRecipient = value;
                OnPropertyChanged(nameof(ToRecipient));
            }
        }

        private bool _isIncludeInternalDomain;
        public bool IsIncludeInternalDomain
        {
            get => _isIncludeInternalDomain;
            set
            {
                _isIncludeInternalDomain = value;
                OnPropertyChanged(nameof(IsIncludeInternalDomain));
            }
        }

        #endregion

        #region AutoAddMessage

        private void LoadAutoAddMessageData()
        {
            var list = CsvFileHandler.ReadCsv<AutoAddMessage>(typeof(AutoAddMessageMap), "AutoAddMessage.csv");
            if (list.Count == 0) return;

            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            _autoAddMessage.AddRange(list);

            //実際に使用するのは1行目の設定のみ
            IsAddToStart = _autoAddMessage[0].IsAddToStart;
            IsAddToEnd = _autoAddMessage[0].IsAddToEnd;
            MessageOfAddToStart = _autoAddMessage[0].MessageOfAddToStart;
            MessageOfAddToEnd = _autoAddMessage[0].MessageOfAddToEnd;
        }

        private async Task SaveAutoAddMessageToCsv()
        {
            var tempAutoAddMessage = new List<AutoAddMessage>
            {
                new AutoAddMessage
                {
                    IsAddToStart = IsAddToStart,
                    IsAddToEnd = IsAddToEnd,
                    MessageOfAddToStart = MessageOfAddToStart,
                    MessageOfAddToEnd = MessageOfAddToEnd
                }
            };

            var list = tempAutoAddMessage.Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AutoAddMessageMap), "AutoAddMessage.csv", list));
        }

        private readonly List<AutoAddMessage> _autoAddMessage = new List<AutoAddMessage>();

        private bool _isAddToStart;
        public bool IsAddToStart
        {
            get => _isAddToStart;
            set
            {
                _isAddToStart = value;
                OnPropertyChanged(nameof(IsAddToStart));
            }
        }

        private bool _isAddToEnd;
        public bool IsAddToEnd
        {
            get => _isAddToEnd;
            set
            {
                _isAddToEnd = value;
                OnPropertyChanged(nameof(IsAddToEnd));
            }
        }

        private string _messageOfAddToStart;
        public string MessageOfAddToStart
        {
            get => _messageOfAddToStart;
            set
            {
                _messageOfAddToStart = value;
                OnPropertyChanged(nameof(MessageOfAddToStart));
            }
        }

        private string _messageOfAddToEnd;
        public string MessageOfAddToEnd
        {
            get => _messageOfAddToEnd;
            set
            {
                _messageOfAddToEnd = value;
                OnPropertyChanged(nameof(MessageOfAddToEnd));
            }
        }

        #endregion

        #region AlertKeywordOfSubjectWhenOpeningMail

        public ICommand ImportAlertKeywordOfSubjectWhenOpeningMailsList { get; }
        public ICommand ExportAlertKeywordOfSubjectWhenOpeningMailsList { get; }

        private void LoadAlertKeywordOfSubjectWhenOpeningMailsData()
        {
            var alertKeywordOfSubjectWhenOpeningMails = CsvFileHandler.ReadCsv<AlertKeywordOfSubjectWhenOpeningMail>(typeof(AlertKeywordOfSubjectWhenOpeningMailMap), "AlertKeywordOfSubjectWhenOpeningMailList.csv");
            foreach (var data in alertKeywordOfSubjectWhenOpeningMails.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
            {
                AlertKeywordOfSubjectWhenOpeningMails.Add(data);
            }
        }

        private async Task SaveAlertKeywordOfSubjectWhenOpeningMailToCsv()
        {
            var list = AlertKeywordOfSubjectWhenOpeningMails.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AlertKeywordOfSubjectWhenOpeningMailMap), "AlertKeywordOfSubjectWhenOpeningMailList.csv", list));
        }

        private void ImportAlertKeywordOfSubjectWhenOpeningMailsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AlertKeywordOfSubjectWhenOpeningMail>(typeof(AlertKeywordOfSubjectWhenOpeningMailMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
                {
                    AlertKeywordOfSubjectWhenOpeningMails.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAlertKeywordOfSubjectWhenOpeningMailsToCsv()
        {
            var list = AlertKeywordOfSubjectWhenOpeningMails.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AlertKeywordOfSubjectWhenOpeningMailMap), list, "AlertKeywordOfSubjectWhenOpeningMailList.csv");
        }

        private ObservableCollection<AlertKeywordOfSubjectWhenOpeningMail> _alertKeywordOfSubjectWhenOpeningMails = new ObservableCollection<AlertKeywordOfSubjectWhenOpeningMail>();
        public ObservableCollection<AlertKeywordOfSubjectWhenOpeningMail> AlertKeywordOfSubjectWhenOpeningMails
        {
            get => _alertKeywordOfSubjectWhenOpeningMails;
            set
            {
                _alertKeywordOfSubjectWhenOpeningMails = value;
                OnPropertyChanged(nameof(AlertKeywordOfSubjectWhenOpeningMails));
            }
        }

        #endregion

        #region AutoDeleteRecipient

        public ICommand ImportAutoDeleteRecipientsList { get; }
        public ICommand ExportAutoDeleteRecipientsList { get; }

        private void LoadAutoDeleteRecipientsData()
        {
            var autoDeleteRecipients = CsvFileHandler.ReadCsv<AutoDeleteRecipient>(typeof(AutoDeleteRecipientMap), "AutoDeleteRecipientList.csv");
            foreach (var data in autoDeleteRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)))
            {
                AutoDeleteRecipients.Add(data);
            }
        }

        private async Task SaveAutoDeleteRecipientToCsv()
        {
            var list = AutoDeleteRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(AutoDeleteRecipientMap), "AutoDeleteRecipientList.csv", list));
        }

        private void ImportAutoDeleteRecipientsFromCsv()
        {
            try
            {
                var importData = CsvFileHandler.ImportCsv<AutoDeleteRecipient>(typeof(AutoDeleteRecipientMap));
                foreach (var data in importData.Where(x => !string.IsNullOrEmpty(x.Recipient)))
                {
                    AutoDeleteRecipients.Add(data);
                }

                _ = MessageBox.Show(Properties.Resources.SuccessfulImport, Properties.Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception)
            {
                _ = MessageBox.Show(Properties.Resources.ImportFailed, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }

        private void ExportAutoDeleteRecipientsToCsv()
        {
            var list = AutoDeleteRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            CsvFileHandler.ExportCsv(typeof(AutoDeleteRecipientMap), list, "AutoDeleteRecipientList.csv");
        }

        private ObservableCollection<AutoDeleteRecipient> _autoDeleteRecipients = new ObservableCollection<AutoDeleteRecipient>();
        public ObservableCollection<AutoDeleteRecipient> AutoDeleteRecipients
        {
            get => _autoDeleteRecipients;
            set
            {
                _autoDeleteRecipients = value;
                OnPropertyChanged(nameof(AutoDeleteRecipient));
            }
        }

        #endregion

        #region SecurityForReceivedMail

        private void LoadSecurityForReceivedMailData()
        {
            var list = CsvFileHandler.ReadCsv<SecurityForReceivedMail>(typeof(SecurityForReceivedMailMap), "SecurityForReceivedMail.csv");
            if (list.Count == 0) return;

            _securityForReceivedMail.AddRange(list);

            //実際に使用するのは1行目の設定のみ
            IsEnableSecurityForReceivedMail = _securityForReceivedMail[0].IsEnableSecurityForReceivedMail;
            IsEnableAlertKeywordOfSubjectWhenOpeningMailsData = _securityForReceivedMail[0].IsEnableAlertKeywordOfSubjectWhenOpeningMailsData;
            IsEnableMailHeaderAnalysis = _securityForReceivedMail[0].IsEnableMailHeaderAnalysis;
            IsShowWarningWhenSpfFails = _securityForReceivedMail[0].IsShowWarningWhenSpfFails;
            IsShowWarningWhenDkimFails = _securityForReceivedMail[0].IsShowWarningWhenDkimFails;
            IsEnableWarningFeatureWhenOpeningAttachments = _securityForReceivedMail[0].IsEnableWarningFeatureWhenOpeningAttachments;
            IsWarnBeforeOpeningAttachments = _securityForReceivedMail[0].IsWarnBeforeOpeningAttachments;
            IsWarnBeforeOpeningEncryptedZip = _securityForReceivedMail[0].IsWarnBeforeOpeningEncryptedZip;
            IsWarnLinkFileInTheZip = _securityForReceivedMail[0].IsWarnLinkFileInTheZip;
            IsWarnOneFileInTheZip = _securityForReceivedMail[0].IsWarnOneFileInTheZip;
            IsWarnOfficeFileWithMacroInTheZip = _securityForReceivedMail[0].IsWarnOfficeFileWithMacroInTheZip;
            IsWarnBeforeOpeningAttachmentsThatContainMacros = _securityForReceivedMail[0].IsWarnBeforeOpeningAttachmentsThatContainMacros;
        }

        private async Task SecurityForReceivedMailToCsv()
        {
            var tempSecurityForReceivedMail = new List<SecurityForReceivedMail>
            {
                new SecurityForReceivedMail
                {
                    IsEnableSecurityForReceivedMail = IsEnableSecurityForReceivedMail,
                    IsEnableAlertKeywordOfSubjectWhenOpeningMailsData = IsEnableAlertKeywordOfSubjectWhenOpeningMailsData,
                    IsEnableMailHeaderAnalysis = IsEnableMailHeaderAnalysis,
                    IsShowWarningWhenSpfFails = IsShowWarningWhenSpfFails,
                    IsShowWarningWhenDkimFails = IsShowWarningWhenDkimFails,
                    IsEnableWarningFeatureWhenOpeningAttachments = IsEnableWarningFeatureWhenOpeningAttachments,
                    IsWarnBeforeOpeningAttachments = IsWarnBeforeOpeningAttachments,
                    IsWarnBeforeOpeningEncryptedZip = IsWarnBeforeOpeningEncryptedZip,
                    IsWarnLinkFileInTheZip = IsWarnLinkFileInTheZip,
                    IsWarnOneFileInTheZip = IsWarnOneFileInTheZip,
                    IsWarnOfficeFileWithMacroInTheZip = IsWarnOfficeFileWithMacroInTheZip,
                    IsWarnBeforeOpeningAttachmentsThatContainMacros = IsWarnBeforeOpeningAttachmentsThatContainMacros
        }
            };

            var list = tempSecurityForReceivedMail.Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(SecurityForReceivedMailMap), "SecurityForReceivedMail.csv", list));
        }

        private readonly List<SecurityForReceivedMail> _securityForReceivedMail = new List<SecurityForReceivedMail>();

        private bool _isEnableSecurityForReceivedMail;
        public bool IsEnableSecurityForReceivedMail
        {
            get => _isEnableSecurityForReceivedMail;
            set
            {
                _isEnableSecurityForReceivedMail = value;
                OnPropertyChanged(nameof(IsEnableSecurityForReceivedMail));
            }
        }

        private bool _isEnableAlertKeywordOfSubjectWhenOpeningMailsData;
        public bool IsEnableAlertKeywordOfSubjectWhenOpeningMailsData
        {
            get => _isEnableAlertKeywordOfSubjectWhenOpeningMailsData;
            set
            {
                _isEnableAlertKeywordOfSubjectWhenOpeningMailsData = value;
                OnPropertyChanged(nameof(IsEnableAlertKeywordOfSubjectWhenOpeningMailsData));
            }
        }

        private bool _isEnableMailHeaderAnalysis;
        public bool IsEnableMailHeaderAnalysis
        {
            get => _isEnableMailHeaderAnalysis;
            set
            {
                _isEnableMailHeaderAnalysis = value;
                OnPropertyChanged(nameof(IsEnableMailHeaderAnalysis));
            }
        }

        private bool _isShowWarningWhenSpfFails;
        public bool IsShowWarningWhenSpfFails
        {
            get => _isShowWarningWhenSpfFails;
            set
            {
                _isShowWarningWhenSpfFails = value;
                OnPropertyChanged(nameof(IsShowWarningWhenSpfFails));
            }
        }

        private bool _isShowWarningWhenDkimFails;
        public bool IsShowWarningWhenDkimFails
        {
            get => _isShowWarningWhenDkimFails;
            set
            {
                _isShowWarningWhenDkimFails = value;
                OnPropertyChanged(nameof(IsShowWarningWhenDkimFails));
            }
        }

        private bool _isEnableWarningFeatureWhenOpeningAttachments;
        public bool IsEnableWarningFeatureWhenOpeningAttachments
        {
            get => _isEnableWarningFeatureWhenOpeningAttachments;
            set
            {
                _isEnableWarningFeatureWhenOpeningAttachments = value;
                OnPropertyChanged(nameof(IsEnableWarningFeatureWhenOpeningAttachments));
            }
        }

        private bool _isWarnBeforeOpeningAttachments;
        public bool IsWarnBeforeOpeningAttachments
        {
            get => _isWarnBeforeOpeningAttachments;
            set
            {
                _isWarnBeforeOpeningAttachments = value;
                OnPropertyChanged(nameof(IsWarnBeforeOpeningAttachments));
            }
        }

        private bool _isWarnBeforeOpeningEncryptedZip;
        public bool IsWarnBeforeOpeningEncryptedZip
        {
            get => _isWarnBeforeOpeningEncryptedZip;
            set
            {
                _isWarnBeforeOpeningEncryptedZip = value;
                OnPropertyChanged(nameof(IsWarnBeforeOpeningEncryptedZip));
            }
        }

        private bool _isWarnLinkFileInTheZip;
        public bool IsWarnLinkFileInTheZip
        {
            get => _isWarnLinkFileInTheZip;
            set
            {
                _isWarnLinkFileInTheZip = value;
                OnPropertyChanged(nameof(IsWarnLinkFileInTheZip));
            }
        }

        private bool _isWarnOneFileInTheZip;
        public bool IsWarnOneFileInTheZip
        {
            get => _isWarnOneFileInTheZip;
            set
            {
                _isWarnOneFileInTheZip = value;
                OnPropertyChanged(nameof(IsWarnOneFileInTheZip));
            }
        }

        private bool _isWarnOfficeFileWithMacroInTheZip;
        public bool IsWarnOfficeFileWithMacroInTheZip
        {
            get => _isWarnOfficeFileWithMacroInTheZip;
            set
            {
                _isWarnOfficeFileWithMacroInTheZip = value;
                OnPropertyChanged(nameof(IsWarnOfficeFileWithMacroInTheZip));
            }
        }

        private bool _isWarnBeforeOpeningAttachmentsThatContainMacros;
        public bool IsWarnBeforeOpeningAttachmentsThatContainMacros
        {
            get => _isWarnBeforeOpeningAttachmentsThatContainMacros;
            set
            {
                _isWarnBeforeOpeningAttachmentsThatContainMacros = value;
                OnPropertyChanged(nameof(IsWarnBeforeOpeningAttachmentsThatContainMacros));
            }
        }

        #endregion  

        #region GeneralSetting

        private void LoadGeneralSettingData()
        {
            var list = CsvFileHandler.ReadCsv<GeneralSetting>(typeof(GeneralSettingMap), "GeneralSetting.csv");
            if (list.Count == 0) return;

            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            _generalSetting.AddRange(list);

            //実際に使用するのは1行目の設定のみ
            IsDoNotConfirmationIfAllRecipientsAreSameDomain = _generalSetting[0].IsDoNotConfirmationIfAllRecipientsAreSameDomain;
            IsDoDoNotConfirmationIfAllWhite = _generalSetting[0].IsDoDoNotConfirmationIfAllWhite;
            IsAutoCheckIfAllRecipientsAreSameDomain = _generalSetting[0].IsAutoCheckIfAllRecipientsAreSameDomain;
            IsShowConfirmationToMultipleDomain = _generalSetting[0].IsShowConfirmationToMultipleDomain;
            EnableForgottenToAttachAlert = _generalSetting[0].EnableForgottenToAttachAlert;
            EnableGetContactGroupMembers = _generalSetting[0].EnableGetContactGroupMembers;
            EnableGetExchangeDistributionListMembers = _generalSetting[0].EnableGetExchangeDistributionListMembers;
            ContactGroupMembersAreWhite = _generalSetting[0].ContactGroupMembersAreWhite;
            ExchangeDistributionListMembersAreWhite = _generalSetting[0].ExchangeDistributionListMembersAreWhite;
            IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = _generalSetting[0].IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles;
            IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain = _generalSetting[0].IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain;
            IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain = _generalSetting[0].IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain;
            IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain = _generalSetting[0].IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain;
            IsEnableRecipientsAreSortedByDomain = _generalSetting[0].IsEnableRecipientsAreSortedByDomain;
            IsAutoAddSenderToBcc = _generalSetting[0].IsAutoAddSenderToBcc;
            IsAutoCheckRegisteredInContacts = _generalSetting[0].IsAutoCheckRegisteredInContacts;
            IsAutoCheckRegisteredInContactsAndMemberOfContactLists = _generalSetting[0].IsAutoCheckRegisteredInContactsAndMemberOfContactLists;
            IsCheckNameAndDomainsFromRecipients = _generalSetting[0].IsCheckNameAndDomainsFromRecipients;
            IsWarningIfRecipientsIsNotRegistered = _generalSetting[0].IsWarningIfRecipientsIsNotRegistered;
            IsProhibitsSendingMailIfRecipientsIsNotRegistered = _generalSetting[0].IsProhibitsSendingMailIfRecipientsIsNotRegistered;
            IsShowConfirmationAtSendMeetingRequest = _generalSetting[0].IsShowConfirmationAtSendMeetingRequest;
            IsAutoAddSenderToCc = _generalSetting[0].IsAutoAddSenderToCc;
            IsCheckNameAndDomainsIncludeSubject = _generalSetting[0].IsCheckNameAndDomainsIncludeSubject;
            IsCheckNameAndDomainsFromSubject = _generalSetting[0].IsCheckNameAndDomainsFromSubject;
            IsShowConfirmationAtSendTaskRequest = _generalSetting[0].IsShowConfirmationAtSendTaskRequest;
            IsAutoCheckAttachments = _generalSetting[0].IsAutoCheckAttachments;
            IsCheckKeywordAndRecipientsIncludeSubject = _generalSetting[0].IsCheckKeywordAndRecipientsIncludeSubject;

            if (_generalSetting[0].LanguageCode is null) return;

            //設定ファイル内に言語設定があればそれをロードする。
            Language.LanguageCode = _generalSetting[0].LanguageCode;
            foreach (var lang in Languages.Where(lang => lang.LanguageCode == Language.LanguageCode))
            {
                LanguageNumber = lang.LanguageNumber;
            }
        }

        private async Task SaveGeneralSettingToCsv()
        {
            var languageCode = Language.LanguageCode ?? CultureInfo.CurrentUICulture.Name;
            if (Language.LanguageCode != null)
            {
                ResourceService.Instance.ChangeCulture(Language.LanguageCode);
            }

            var tempGeneralSetting = new List<GeneralSetting>
            {
                new GeneralSetting
                {
                    IsDoNotConfirmationIfAllRecipientsAreSameDomain = IsDoNotConfirmationIfAllRecipientsAreSameDomain,
                    IsDoDoNotConfirmationIfAllWhite = IsDoDoNotConfirmationIfAllWhite,
                    IsAutoCheckIfAllRecipientsAreSameDomain = IsAutoCheckIfAllRecipientsAreSameDomain,
                    LanguageCode = languageCode,
                    IsShowConfirmationToMultipleDomain = IsShowConfirmationToMultipleDomain,
                    EnableForgottenToAttachAlert = EnableForgottenToAttachAlert,
                    EnableGetContactGroupMembers = EnableGetContactGroupMembers,
                    EnableGetExchangeDistributionListMembers = EnableGetExchangeDistributionListMembers,
                    ContactGroupMembersAreWhite = ContactGroupMembersAreWhite,
                    ExchangeDistributionListMembersAreWhite = ExchangeDistributionListMembersAreWhite,
                    IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles,
                    IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain = IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain,
                    IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain = IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain,
                    IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain = IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain,
                    IsEnableRecipientsAreSortedByDomain = IsEnableRecipientsAreSortedByDomain,
                    IsAutoAddSenderToBcc = IsAutoAddSenderToBcc,
                    IsAutoCheckRegisteredInContacts = IsAutoCheckRegisteredInContacts,
                    IsAutoCheckRegisteredInContactsAndMemberOfContactLists = IsAutoCheckRegisteredInContactsAndMemberOfContactLists,
                    IsCheckNameAndDomainsFromRecipients = IsCheckNameAndDomainsFromRecipients,
                    IsWarningIfRecipientsIsNotRegistered = IsWarningIfRecipientsIsNotRegistered,
                    IsProhibitsSendingMailIfRecipientsIsNotRegistered = IsProhibitsSendingMailIfRecipientsIsNotRegistered,
                    IsShowConfirmationAtSendMeetingRequest = IsShowConfirmationAtSendMeetingRequest,
                    IsAutoAddSenderToCc = IsAutoAddSenderToCc,
                    IsCheckNameAndDomainsIncludeSubject = IsCheckNameAndDomainsIncludeSubject,
                    IsCheckNameAndDomainsFromSubject =IsCheckNameAndDomainsFromSubject,
                    IsShowConfirmationAtSendTaskRequest = IsShowConfirmationAtSendTaskRequest,
                    IsAutoCheckAttachments = IsAutoCheckAttachments,
                    IsCheckKeywordAndRecipientsIncludeSubject = IsCheckKeywordAndRecipientsIncludeSubject
                }
            };

            var list = tempGeneralSetting.Cast<object>().ToList();
            await Task.Run(() => CsvFileHandler.CreateOrReplaceCsv(typeof(GeneralSettingMap), "GeneralSetting.csv", list));
        }

        private readonly List<GeneralSetting> _generalSetting = new List<GeneralSetting>();

        private bool _isDoNotConfirmationIfAllRecipientsAreSameDomain;
        public bool IsDoNotConfirmationIfAllRecipientsAreSameDomain
        {
            get => _isDoNotConfirmationIfAllRecipientsAreSameDomain;
            set
            {
                _isDoNotConfirmationIfAllRecipientsAreSameDomain = value;
                OnPropertyChanged(nameof(IsDoNotConfirmationIfAllRecipientsAreSameDomain));
            }
        }

        private bool _isDoDoNotConfirmationIfAllWhite;
        public bool IsDoDoNotConfirmationIfAllWhite
        {
            get => _isDoDoNotConfirmationIfAllWhite;
            set
            {
                _isDoDoNotConfirmationIfAllWhite = value;
                OnPropertyChanged(nameof(IsDoDoNotConfirmationIfAllWhite));
            }
        }

        private bool _isAutoCheckIfAllRecipientsAreSameDomain;
        public bool IsAutoCheckIfAllRecipientsAreSameDomain
        {
            get => _isAutoCheckIfAllRecipientsAreSameDomain;
            set
            {
                _isAutoCheckIfAllRecipientsAreSameDomain = value;
                OnPropertyChanged(nameof(IsAutoCheckIfAllRecipientsAreSameDomain));
            }
        }

        private bool _isShowConfirmationToMultipleDomain;
        public bool IsShowConfirmationToMultipleDomain
        {
            get => _isShowConfirmationToMultipleDomain;
            set
            {
                _isShowConfirmationToMultipleDomain = value;
                OnPropertyChanged(nameof(IsShowConfirmationToMultipleDomain));
            }
        }

        private bool _enableForgottenToAttachAlert = true;
        public bool EnableForgottenToAttachAlert
        {
            get => _enableForgottenToAttachAlert;
            set
            {
                _enableForgottenToAttachAlert = value;
                OnPropertyChanged(nameof(EnableForgottenToAttachAlert));
            }
        }

        private bool _enableGetContactGroupMembers;
        public bool EnableGetContactGroupMembers
        {
            get => _enableGetContactGroupMembers;
            set
            {
                _enableGetContactGroupMembers = value;
                OnPropertyChanged(nameof(EnableGetContactGroupMembers));
            }
        }

        private bool _enableGetExchangeDistributionListMembers;
        public bool EnableGetExchangeDistributionListMembers
        {
            get => _enableGetExchangeDistributionListMembers;
            set
            {
                _enableGetExchangeDistributionListMembers = value;
                OnPropertyChanged(nameof(EnableGetExchangeDistributionListMembers));
            }
        }

        private bool _contactGroupMembersAreWhite = true;
        public bool ContactGroupMembersAreWhite
        {
            get => _contactGroupMembersAreWhite;
            set
            {
                _contactGroupMembersAreWhite = value;
                OnPropertyChanged(nameof(ContactGroupMembersAreWhite));
            }
        }

        private bool _exchangeDistributionListMembersAreWhite = true;
        public bool ExchangeDistributionListMembersAreWhite
        {
            get => _exchangeDistributionListMembersAreWhite;
            set
            {
                _exchangeDistributionListMembersAreWhite = value;
                OnPropertyChanged(nameof(ExchangeDistributionListMembersAreWhite));
            }
        }

        private bool _isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles;
        public bool IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles
        {
            get => _isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles;
            set
            {
                _isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = value;
                OnPropertyChanged(nameof(IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles));
            }
        }

        private bool _isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain;
        public bool IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain
        {
            get => _isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain;
            set
            {
                _isDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain = value;
                OnPropertyChanged(nameof(IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain));
            }
        }

        private bool _isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain;
        public bool IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain
        {
            get => _isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain;
            set
            {
                _isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain = value;
                OnPropertyChanged(nameof(IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain));
            }
        }

        private bool _isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain;
        public bool IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain
        {
            get => _isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain;
            set
            {
                _isDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain = value;
                OnPropertyChanged(nameof(IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain));
            }
        }

        private bool _isEnableRecipientsAreSortedByDomain;
        public bool IsEnableRecipientsAreSortedByDomain
        {
            get => _isEnableRecipientsAreSortedByDomain;
            set
            {
                _isEnableRecipientsAreSortedByDomain = value;
                OnPropertyChanged(nameof(IsEnableRecipientsAreSortedByDomain));
            }
        }

        private bool _isAutoAddSenderToBcc;
        public bool IsAutoAddSenderToBcc
        {
            get => _isAutoAddSenderToBcc;
            set
            {
                _isAutoAddSenderToBcc = value;
                OnPropertyChanged(nameof(IsAutoAddSenderToBcc));
            }
        }

        private bool _isAutoCheckRegisteredInContacts;
        public bool IsAutoCheckRegisteredInContacts
        {
            get => _isAutoCheckRegisteredInContacts;
            set
            {
                _isAutoCheckRegisteredInContacts = value;
                OnPropertyChanged(nameof(IsAutoCheckRegisteredInContacts));
            }
        }

        private bool _isAutoCheckRegisteredInContactsAndMemberOfContactLists;
        public bool IsAutoCheckRegisteredInContactsAndMemberOfContactLists
        {
            get => _isAutoCheckRegisteredInContactsAndMemberOfContactLists;
            set
            {
                _isAutoCheckRegisteredInContactsAndMemberOfContactLists = value;
                OnPropertyChanged(nameof(IsAutoCheckRegisteredInContactsAndMemberOfContactLists));
            }
        }

        private bool _isCheckNameAndDomainsFromRecipients;
        public bool IsCheckNameAndDomainsFromRecipients
        {
            get => _isCheckNameAndDomainsFromRecipients;
            set
            {
                _isCheckNameAndDomainsFromRecipients = value;
                OnPropertyChanged(nameof(IsCheckNameAndDomainsFromRecipients));
            }
        }

        private bool _isWarningIfRecipientsIsNotRegistered;
        public bool IsWarningIfRecipientsIsNotRegistered
        {
            get => _isWarningIfRecipientsIsNotRegistered;
            set
            {
                _isWarningIfRecipientsIsNotRegistered = value;
                OnPropertyChanged(nameof(IsWarningIfRecipientsIsNotRegistered));
            }
        }

        private bool _isProhibitsSendingMailIfRecipientsIsNotRegistered;
        public bool IsProhibitsSendingMailIfRecipientsIsNotRegistered
        {
            get => _isProhibitsSendingMailIfRecipientsIsNotRegistered;
            set
            {
                _isProhibitsSendingMailIfRecipientsIsNotRegistered = value;
                OnPropertyChanged(nameof(IsProhibitsSendingMailIfRecipientsIsNotRegistered));
                OnPropertyChanged(nameof(IsWarningIfRecipientsIsNotRegisteredCheckBoxIsEnabled));
            }
        }

        private bool _isShowConfirmationAtSendMeetingRequest;
        public bool IsShowConfirmationAtSendMeetingRequest
        {
            get => _isShowConfirmationAtSendMeetingRequest;
            set
            {
                _isShowConfirmationAtSendMeetingRequest = value;
                OnPropertyChanged(nameof(IsShowConfirmationAtSendMeetingRequest));
            }
        }

        private bool _isAutoAddSenderToCc;
        public bool IsAutoAddSenderToCc
        {
            get => _isAutoAddSenderToCc;
            set
            {
                _isAutoAddSenderToCc = value;
                OnPropertyChanged(nameof(IsAutoAddSenderToCc));
            }
        }

        private bool _isCheckNameAndDomainsIncludeSubject;
        public bool IsCheckNameAndDomainsIncludeSubject
        {
            get => _isCheckNameAndDomainsIncludeSubject;
            set
            {
                _isCheckNameAndDomainsIncludeSubject = value;
                OnPropertyChanged(nameof(IsCheckNameAndDomainsIncludeSubject));
            }
        }

        private bool _isCheckNameAndDomainsFromSubject;
        public bool IsCheckNameAndDomainsFromSubject
        {
            get => _isCheckNameAndDomainsFromSubject;
            set
            {
                _isCheckNameAndDomainsFromSubject = value;
                OnPropertyChanged(nameof(IsCheckNameAndDomainsFromSubject));
            }
        }

        private bool _isShowConfirmationAtSendTaskRequest;
        public bool IsShowConfirmationAtSendTaskRequest
        {
            get => _isShowConfirmationAtSendTaskRequest;
            set
            {
                _isShowConfirmationAtSendTaskRequest = value;
                OnPropertyChanged(nameof(IsShowConfirmationAtSendTaskRequest));
            }
        }

        private bool _isAutoCheckAttachments;
        public bool IsAutoCheckAttachments
        {
            get => _isAutoCheckAttachments;
            set
            {
                _isAutoCheckAttachments = value;
                OnPropertyChanged(nameof(IsAutoCheckAttachments));
            }
        }

        private bool _isCheckKeywordAndRecipientsIncludeSubject;
        public bool IsCheckKeywordAndRecipientsIncludeSubject
        {
            get => _isCheckKeywordAndRecipientsIncludeSubject;
            set
            {
                _isCheckKeywordAndRecipientsIncludeSubject = value;
                OnPropertyChanged(nameof(IsCheckKeywordAndRecipientsIncludeSubject));
            }
        }

        public bool IsWarningIfRecipientsIsNotRegisteredCheckBoxIsEnabled => !IsProhibitsSendingMailIfRecipientsIsNotRegistered;

        private LanguageCodeAndName _language = new LanguageCodeAndName();
        public LanguageCodeAndName Language
        {
            get => _language;
            set
            {
                _language = value;
                OnPropertyChanged(nameof(Language));
            }
        }

        private List<LanguageCodeAndName> _languages;
        public List<LanguageCodeAndName> Languages
        {
            get => _languages;
            set
            {
                _languages = value;
                OnPropertyChanged(nameof(Languages));
            }
        }

        private int _languageNumber = -1;
        public int LanguageNumber
        {
            get => _languageNumber;
            set
            {
                _languageNumber = value;
                OnPropertyChanged(nameof(LanguageNumber));
            }
        }

        #endregion

    }
}