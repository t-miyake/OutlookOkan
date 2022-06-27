using OutlookOkan.CsvTools;
using OutlookOkan.Services;
using OutlookOkan.Types;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

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

            //Load language code and name.
            var languages = new Languages();
            Languages = languages.Language;

            //Load settings from csv.
            LoadGeneralSettingData();
            LoadWhitelistData();
            LoadNameAndDomainsData();
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
            LoadAutoAddMessageData();
        }

        internal async Task SaveSettings()
        {
            IEnumerable<Task> saveTasks = new[]
                {
                    SaveGeneralSettingToCsv(),
                    SaveWhiteListToCsv(),
                    SaveNameAndDomainsToCsv(),
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
                    SaveAutoAddMessageToCsv()
                };

            await Task.WhenAll(saveTasks);
        }

        #region Whitelist

        public ICommand ImportWhiteList { get; }
        public ICommand ExportWhiteList { get; }

        private void LoadWhitelistData()
        {
            var readCsv = new ReadAndWriteCsv("Whitelist.csv");
            var whitelist = readCsv.GetCsvRecords<Whitelist>(readCsv.LoadCsv<WhitelistMap>());

            foreach (var data in whitelist.Where(x => !string.IsNullOrEmpty(x.WhiteName)))
            {
                Whitelist.Add(data);
            }
        }

        private async Task SaveWhiteListToCsv()
        {
            var list = Whitelist.Where(x => !string.IsNullOrEmpty(x.WhiteName)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("Whitelist.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<WhitelistMap>(list));
        }

        private void ImportWhiteListFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<Whitelist>(importAction.GetCsvRecords<Whitelist>(importAction.LoadCsv<WhitelistMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<WhitelistMap>(list, "Whitelist.csv");
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
            var readCsv = new ReadAndWriteCsv("NameAndDomains.csv");
            var nameAndDomains = readCsv.GetCsvRecords<NameAndDomains>(readCsv.LoadCsv<NameAndDomainsMap>());

            foreach (var data in nameAndDomains.Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)))
            {
                NameAndDomains.Add(data);
            }
        }

        private async Task SaveNameAndDomainsToCsv()
        {
            var list = NameAndDomains.Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("NameAndDomains.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<NameAndDomainsMap>(list));
        }

        private void ImportNameAndDomainsFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<NameAndDomains>(importAction.GetCsvRecords<NameAndDomains>(importAction.LoadCsv<NameAndDomainsMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<NameAndDomainsMap>(list, "NameAndDomains.csv");
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

        #region AlertKeywordAndMessagesForSubject

        public ICommand ImportAlertKeywordAndMessagesForSubjectList { get; }
        public ICommand ExportAlertKeywordAndMessagesForSubjectList { get; }

        private void LoadAlertKeywordAndMessagesForSubjectData()
        {
            var readCsv = new ReadAndWriteCsv("AlertKeywordAndMessageListForSubject.csv");
            var alertKeywordAndMessagesForSubject = readCsv.GetCsvRecords<AlertKeywordAndMessageForSubject>(readCsv.LoadCsv<AlertKeywordAndMessageForSubjectMap>());

            foreach (var data in alertKeywordAndMessagesForSubject.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
            {
                AlertKeywordAndMessagesForSubject.Add(data);
            }
        }

        private async Task SaveAlertKeywordAndMessageForSubjectToCsv()
        {
            var list = AlertKeywordAndMessagesForSubject.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AlertKeywordAndMessageListForSubject.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AlertKeywordAndMessageForSubjectMap>(list));
        }

        private void ImportAlertKeywordAndMessagesForSubjectFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AlertKeywordAndMessageForSubject>(importAction.GetCsvRecords<AlertKeywordAndMessageForSubject>(importAction.LoadCsv<AlertKeywordAndMessageForSubjectMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AlertKeywordAndMessageForSubjectMap>(list, "AlertKeywordAndMessageListForSubject.csv");
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
            var readCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            var alertKeywordAndMessages = readCsv.GetCsvRecords<AlertKeywordAndMessage>(readCsv.LoadCsv<AlertKeywordAndMessageMap>());

            foreach (var data in alertKeywordAndMessages.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)))
            {
                AlertKeywordAndMessages.Add(data);
            }
        }

        private async Task SaveAlertKeywordAndMessageToCsv()
        {
            var list = AlertKeywordAndMessages.Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AlertKeywordAndMessageMap>(list));
        }

        private void ImportAlertKeywordAndMessagesFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AlertKeywordAndMessage>(importAction.GetCsvRecords<AlertKeywordAndMessage>(importAction.LoadCsv<AlertKeywordAndMessageMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AlertKeywordAndMessageMap>(list, "AlertKeywordAndMessageList.csv");
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
            var readCsv = new ReadAndWriteCsv("AlertAddressList.csv");
            var alertAddresses = readCsv.GetCsvRecords<AlertAddress>(readCsv.LoadCsv<AlertAddressMap>());

            foreach (var data in alertAddresses.Where(x => !string.IsNullOrEmpty(x.TargetAddress)))
            {
                AlertAddresses.Add(data);
            }
        }

        private async Task SaveAlertAddressesToCsv()
        {
            var list = AlertAddresses.Where(x => !string.IsNullOrEmpty(x.TargetAddress)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AlertAddressList.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AlertAddressMap>(list));
        }

        private void ImportAlertAddressesFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AlertAddress>(importAction.GetCsvRecords<AlertAddress>(importAction.LoadCsv<AlertAddressMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AlertAddressMap>(list, "AlertAddressList.csv");
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
            var readCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");
            var autoCcBccKeywords = readCsv.GetCsvRecords<AutoCcBccKeyword>(readCsv.LoadCsv<AutoCcBccKeywordMap>());

            foreach (var data in autoCcBccKeywords.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.AutoAddAddress)))
            {
                AutoCcBccKeywords.Add(data);
            }
        }

        private async Task SaveAutoCcBccKeywordsToCsv()
        {
            var list = AutoCcBccKeywords.Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AutoCcBccKeywordMap>(list));
        }

        private void ImportAutoCcBccKeywordsFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AutoCcBccKeyword>(importAction.GetCsvRecords<AutoCcBccKeyword>(importAction.LoadCsv<AutoCcBccKeywordMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AutoCcBccKeywordMap>(list, "AutoCcBccKeywordList.csv");
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
            var readCsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");
            var autoCcBccRecipient = readCsv.GetCsvRecords<AutoCcBccRecipient>(readCsv.LoadCsv<AutoCcBccRecipientMap>());

            foreach (var data in autoCcBccRecipient.Where(x => !string.IsNullOrEmpty(x.TargetRecipient) && !string.IsNullOrEmpty(x.AutoAddAddress)))
            {
                AutoCcBccRecipients.Add(data);
            }
        }

        private async Task SaveAutoCcBccRecipientsToCsv()
        {
            var list = AutoCcBccRecipients.Where(x => !string.IsNullOrEmpty(x.TargetRecipient) && !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AutoCcBccRecipientMap>(list));
        }

        private void ImportAutoCcBccRecipientsFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AutoCcBccRecipient>(importAction.GetCsvRecords<AutoCcBccRecipient>(importAction.LoadCsv<AutoCcBccRecipientMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AutoCcBccRecipientMap>(list, "AutoCcBccRecipientList.csv");
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
            var readCsv = new ReadAndWriteCsv("AutoCcBccAttachedFileList.csv");
            var autoCcBccAttachedFile = readCsv.GetCsvRecords<AutoCcBccAttachedFile>(readCsv.LoadCsv<AutoCcBccAttachedFileMap>());

            foreach (var data in autoCcBccAttachedFile.Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)))
            {
                AutoCcBccAttachedFiles.Add(data);
            }
        }

        private async Task SaveAutoCcBccAttachedFilesToCsv()
        {
            var list = AutoCcBccAttachedFiles.Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AutoCcBccAttachedFileList.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AutoCcBccAttachedFileMap>(list));
        }

        private void ImportAutoCcBccAttachedFilesFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AutoCcBccAttachedFile>(importAction.GetCsvRecords<AutoCcBccAttachedFile>(importAction.LoadCsv<AutoCcBccAttachedFileMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AutoCcBccAttachedFileMap>(list, "AutoCcBccAttachedFileList.csv");
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
            var readCsv = new ReadAndWriteCsv("DeferredDeliveryMinutes.csv");
            var deferredDeliveryMinutes = readCsv.GetCsvRecords<DeferredDeliveryMinutes>(readCsv.LoadCsv<DeferredDeliveryMinutesMap>());

            foreach (var data in deferredDeliveryMinutes.Where(x => !string.IsNullOrEmpty(x.TargetAddress)))
            {
                DeferredDeliveryMinutes.Add(data);
            }
        }

        private async Task SaveDeferredDeliveryMinutesToCsv()
        {
            var list = DeferredDeliveryMinutes.Where(x => !string.IsNullOrEmpty(x.TargetAddress)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("DeferredDeliveryMinutes.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<DeferredDeliveryMinutesMap>(list));
        }

        private void ImportDeferredDeliveryMinutesFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<DeferredDeliveryMinutes>(importAction.GetCsvRecords<DeferredDeliveryMinutes>(importAction.LoadCsv<DeferredDeliveryMinutesMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<DeferredDeliveryMinutesMap>(list, "DeferredDeliveryMinutes.csv");
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
            var readCsv = new ReadAndWriteCsv("InternalDomainList.csv");
            var internalDomainList = readCsv.GetCsvRecords<InternalDomain>(readCsv.LoadCsv<InternalDomainMap>());

            foreach (var data in internalDomainList.Where(x => !string.IsNullOrEmpty(x.Domain)))
            {
                InternalDomainList.Add(data);
            }
        }

        private async Task SaveInternalDomainListToCsv()
        {
            var list = InternalDomainList.Where(x => !string.IsNullOrEmpty(x.Domain)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("InternalDomainList.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<InternalDomainMap>(list));
        }

        private void ImportInternalDomainListFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<InternalDomain>(importAction.GetCsvRecords<InternalDomain>(importAction.LoadCsv<InternalDomainMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<InternalDomainMap>(list, "InternalDomainList.csv");
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
            var readCsv = new ReadAndWriteCsv("ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            foreach (var data in readCsv.GetCsvRecords<ExternalDomainsWarningAndAutoChangeToBcc>(readCsv.LoadCsv<ExternalDomainsWarningAndAutoChangeToBccMap>()))
            {
                _externalDomainsWarningAndAutoChangeToBcc.Add(data);
            }

            if (_externalDomainsWarningAndAutoChangeToBcc.Count == 0) return;

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
            var writeCsv = new ReadAndWriteCsv("ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<ExternalDomainsWarningAndAutoChangeToBccMap>(list));
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
            var readCsv = new ReadAndWriteCsv("AttachmentsSetting.csv");
            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            foreach (var data in readCsv.GetCsvRecords<AttachmentsSetting>(readCsv.LoadCsv<AttachmentsSettingMap>()))
            {
                _attachmentsSetting.Add(data);
            }

            if (_attachmentsSetting.Count == 0) return;

            //実際に使用するのは1行目の設定のみ
            IsWarningWhenEncryptedZipIsAttached = _attachmentsSetting[0].IsWarningWhenEncryptedZipIsAttached;
            IsProhibitedWhenEncryptedZipIsAttached = _attachmentsSetting[0].IsProhibitedWhenEncryptedZipIsAttached;
            IsEnableAllAttachedFilesAreDetectEncryptedZip = _attachmentsSetting[0].IsEnableAllAttachedFilesAreDetectEncryptedZip;
            IsAttachmentsProhibited = _attachmentsSetting[0].IsAttachmentsProhibited;
            IsWarningWhenAttachedRealFile = _attachmentsSetting[0].IsWarningWhenAttachedRealFile;
            IsEnableOpenAttachedFiles = _attachmentsSetting[0].IsEnableOpenAttachedFiles;
            TargetAttachmentFileExtensionOfOpen = _attachmentsSetting[0].TargetAttachmentFileExtensionOfOpen;
            IsMustOpenBeforeCheckTheAttachedFiles = _attachmentsSetting[0].IsMustOpenBeforeCheckTheAttachedFiles;
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
                    IsMustOpenBeforeCheckTheAttachedFiles = IsMustOpenBeforeCheckTheAttachedFiles

                }
            };

            var list = tempAttachmentsSetting.Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AttachmentsSetting.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AttachmentsSettingMap>(list));
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

        private string _targetAttachmentFileExtensionOfOpen = ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.fif,.pub,.vsd,.vsdx";
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
            var readCsv = new ReadAndWriteCsv("RecipientsAndAttachmentsName.csv");
            var recipientsAndAttachmentsName = readCsv.GetCsvRecords<RecipientsAndAttachmentsName>(readCsv.LoadCsv<RecipientsAndAttachmentsNameMap>());

            foreach (var data in recipientsAndAttachmentsName.Where(x => !string.IsNullOrEmpty(x.AttachmentsName) && !string.IsNullOrEmpty(x.Recipient)))
            {
                RecipientsAndAttachmentsName.Add(data);
            }
        }

        private async Task SaveRecipientsAndAttachmentsNameToCsv()
        {
            var list = RecipientsAndAttachmentsName.Where(x => !string.IsNullOrEmpty(x.AttachmentsName) && !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("RecipientsAndAttachmentsName.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<RecipientsAndAttachmentsNameMap>(list));
        }

        private void ImportRecipientsAndAttachmentsNameFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<RecipientsAndAttachmentsName>(importAction.GetCsvRecords<RecipientsAndAttachmentsName>(importAction.LoadCsv<RecipientsAndAttachmentsNameMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<RecipientsAndAttachmentsNameMap>(list, "RecipientsAndAttachmentsName.csv");
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
            var readCsv = new ReadAndWriteCsv("AttachmentProhibitedRecipients.csv");
            var attachmentProhibitedRecipients = readCsv.GetCsvRecords<AttachmentProhibitedRecipients>(readCsv.LoadCsv<AttachmentProhibitedRecipientsMap>());

            foreach (var data in attachmentProhibitedRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)))
            {
                AttachmentProhibitedRecipients.Add(data);
            }
        }

        private async Task SaveAttachmentProhibitedRecipientsToCsv()
        {
            var list = AttachmentProhibitedRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AttachmentProhibitedRecipients.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AttachmentProhibitedRecipientsMap>(list));
        }

        private void ImportAttachmentProhibitedRecipientsFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AttachmentProhibitedRecipients>(importAction.GetCsvRecords<AttachmentProhibitedRecipients>(importAction.LoadCsv<AttachmentProhibitedRecipientsMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AttachmentProhibitedRecipientsMap>(list, "AttachmentProhibitedRecipients.csv");
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
            var readCsv = new ReadAndWriteCsv("AttachmentAlertRecipients.csv");
            var attachmentAlertRecipients = readCsv.GetCsvRecords<AttachmentAlertRecipients>(readCsv.LoadCsv<AttachmentAlertRecipientsMap>());

            foreach (var data in attachmentAlertRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)))
            {
                AttachmentAlertRecipients.Add(data);
            }
        }

        private async Task SaveAttachmentAlertRecipientsToCsv()
        {
            var list = AttachmentAlertRecipients.Where(x => !string.IsNullOrEmpty(x.Recipient)).Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("AttachmentAlertRecipients.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AttachmentAlertRecipientsMap>(list));
        }

        private void ImportAttachmentAlertRecipientsFromCsv()
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath is null) return;

            try
            {
                var importData = new List<AttachmentAlertRecipients>(importAction.GetCsvRecords<AttachmentAlertRecipients>(importAction.LoadCsv<AttachmentAlertRecipientsMap>(filePath)));
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
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AttachmentAlertRecipientsMap>(list, "AttachmentAlertRecipients.csv");
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
            var readCsv = new ReadAndWriteCsv("ForceAutoChangeRecipientsToBcc.csv");
            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            foreach (var data in readCsv.GetCsvRecords<ForceAutoChangeRecipientsToBcc>(readCsv.LoadCsv<ForceAutoChangeRecipientsToBccMap>()))
            {
                _forceAutoChangeRecipientsToBcc.Add(data);
            }

            if (_forceAutoChangeRecipientsToBcc.Count == 0) return;

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
            var writeCsv = new ReadAndWriteCsv("ForceAutoChangeRecipientsToBcc.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<ForceAutoChangeRecipientsToBccMap>(list));
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
            var readCsv = new ReadAndWriteCsv("AutoAddMessage.csv");
            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            foreach (var data in readCsv.GetCsvRecords<AutoAddMessage>(readCsv.LoadCsv<AutoAddMessageMap>()))
            {
                _autoAddMessage.Add(data);
            }

            if (_autoAddMessage.Count == 0) return;

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
            var writeCsv = new ReadAndWriteCsv("AutoAddMessage.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<AutoAddMessageMap>(list));
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

        #region GeneralSetting

        private void LoadGeneralSettingData()
        {
            var readCsv = new ReadAndWriteCsv("GeneralSetting.csv");
            //1行しかないはずだが、2行以上あるとロード時にエラーとなる恐れがあるため、全行ロードする。
            foreach (var data in readCsv.GetCsvRecords<GeneralSetting>(readCsv.LoadCsv<GeneralSettingMap>()))
            {
                _generalSetting.Add(data);
            }

            if (_generalSetting.Count == 0) return;

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
                    IsCheckNameAndDomainsFromSubject =IsCheckNameAndDomainsFromSubject
                }
            };

            var list = tempGeneralSetting.Cast<object>().ToList();
            var writeCsv = new ReadAndWriteCsv("GeneralSetting.csv");
            await Task.Run(() => writeCsv.WriteRecordsToCsv<GeneralSettingMap>(list));
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