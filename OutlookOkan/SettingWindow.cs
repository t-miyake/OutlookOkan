using System;
using System.Windows.Forms;

namespace OutlookOkan
{
    //TODO 全体を簡潔に修正する。
    //TODO 入力規則を付ける。
    //TODO CSVインポート時のフォーマットチェック。
    public partial class SettingWindow : Form
    {
        public SettingWindow()
        {
            InitializeComponent();

            //Load settings.
            WhitelistToGrid();
            SetNameAndDomainsListToGrid();
            AlertKeywordAndMessageListToGrid();
            AlertAddressListToGrid();
            AutoCcBccKeywordListToGrid();
            AutoCcBccRecipientListToGrid();
        }

        #region BindableLists
        public BindingSource BindableWhitelist { get; set; }
        public BindingSource BindableNameAdnDomainList { get; set; }
        public BindingSource BindableAlertKeywordAndMessageList { get; set; }
        public BindingSource BindableAlertAddressList { get; set; }
        public BindingSource BindableAutoCcBccKeywordList { get; set; }
        public BindingSource BindableAutoCcBccRecipientList { get; set; }
        #endregion

        #region Whitelist setting
        public void WhitelistToGrid()
        {
            var readCsv = new ReadAndWriteCsv("Whitelist.csv");

            BindableWhitelist = new BindingSource(readCsv.ReadCsv<Whitelist>(readCsv.ParseCsv<WhitelistMap>()), string.Empty);
            WhitelistGrid.DataSource = BindableWhitelist;

            WhitelistGrid.Columns[0].HeaderText = @"アドレスまたはドメイン";
        }

        public void SaveWhitelistToCsv()
        {
            var writeCsv = new ReadAndWriteCsv("Whitelist.csv");
            writeCsv.WriteBindableDataToCsv<WhitelistMap>(BindableWhitelist);
        }

        private void WhitelistCsvImportButton_Click(object sender, EventArgs e)
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath != null)
            {
                try
                {
                    var importData = new BindingSource(
                        importAction.ReadCsv<Whitelist>(importAction.ParseCsv<WhitelistMap>(filePath)), string.Empty);

                    foreach (var data in importData)
                    {
                        BindableWhitelist.Add(data);
                    }

                    MessageBox.Show("インポートが完了しました。");
                }
                catch (Exception)
                {
                    MessageBox.Show("インポートに失敗しました。");
                }
            }
        }

        private void WhitelistCsvExportButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<WhitelistMap>(BindableWhitelist, "ホワイトリスト.csv");
        }
        #endregion

        #region NameAndDomainsList setting
        public void SetNameAndDomainsListToGrid()
        {            
            var readCsv = new ReadAndWriteCsv("NameAndDomains.csv");

            BindableNameAdnDomainList = new BindingSource(readCsv.ReadCsv<NameAndDomains>(readCsv.ParseCsv<NameAndDomainsMap>()), string.Empty);
            NameAndDomainsGrid.DataSource = BindableNameAdnDomainList;

            NameAndDomainsGrid.Columns[0].HeaderText = @"名称";
            NameAndDomainsGrid.Columns[1].HeaderText = @"ドメイン (@から)";
        }

        public void SaveNameAndDomainsListToCsv()
        {
            var writeCsv = new ReadAndWriteCsv("NameAndDomains.csv");
            writeCsv.WriteBindableDataToCsv<NameAndDomainsMap>(BindableNameAdnDomainList);
        }

        private void NameAndDomainsCsvImportButton_Click(object sender, EventArgs e)
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath != null)
            {
                try
                {
                    var importData = new BindingSource(
                        importAction.ReadCsv<NameAndDomains>(importAction.ParseCsv<NameAndDomainsMap>(filePath)),
                        string.Empty);
                    foreach (var data in importData)
                    {
                        BindableNameAdnDomainList.Add(data);
                    }

                    MessageBox.Show("インポートが完了しました。");
                }
                catch (Exception)
                {
                    MessageBox.Show("インポートに失敗しました。");
                }
            }
        }

        private void NameAndDomainsCsvExportButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<NameAndDomainsMap>(BindableNameAdnDomainList,"名称とドメインのリスト.csv");
        }
        #endregion

        #region AlertKeywordAndMessageList setting
        public void AlertKeywordAndMessageListToGrid()
        {
            var readCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");

            BindableAlertKeywordAndMessageList = new BindingSource(readCsv.ReadCsv<AlertKeywordAndMessage>(readCsv.ParseCsv<AlertKeywordAndMessageMap>()), string.Empty);
            AlertKeywordAndMessageGrid.DataSource = BindableAlertKeywordAndMessageList;

            AlertKeywordAndMessageGrid.Columns[0].HeaderText = @"警告するキーワード";
            AlertKeywordAndMessageGrid.Columns[1].HeaderText = @"警告文";
        }

        public void SaveAlertKeywordAndMessageListToCsv()
        {
            var writeCsv = new ReadAndWriteCsv("AlertKeywordAndMessageList.csv");
            writeCsv.WriteBindableDataToCsv<AlertKeywordAndMessageMap>(BindableAlertKeywordAndMessageList);
        }

        private void AlertKeywordAndMessageCsvImportButton_Click(object sender, EventArgs e)
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath != null)
            {
                try
                {
                    var importData = new BindingSource(
                        importAction.ReadCsv<AlertKeywordAndMessage>(
                            importAction.ParseCsv<AlertKeywordAndMessageMap>(filePath)), string.Empty);
                    foreach (var data in importData)
                    {
                        BindableAlertKeywordAndMessageList.Add(data);
                    }

                    MessageBox.Show("インポートが完了しました。");
                }
                catch (Exception)
                {
                    MessageBox.Show("インポートに失敗しました。");
                }
            }
        }

        private void AlertKeywordAndMessageCsvExportButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AlertKeywordAndMessageMap>(BindableAlertKeywordAndMessageList, "警告キーワードと警告文.csv");
        }
        #endregion

        #region AlertAddressList setting
        public void AlertAddressListToGrid()
        {
            var readCsv = new ReadAndWriteCsv("AlertAddressList.csv");

            BindableAlertAddressList = new BindingSource(readCsv.ReadCsv<AlertAddress>(readCsv.ParseCsv<AlertAddressMap>()), string.Empty);
            AlertAddressGrid.DataSource = BindableAlertAddressList;

            AlertAddressGrid.Columns[0].HeaderText = @"警告するアドレスまたはドメイン";
        }

        public void SaveAlertAddressListToCsv()
        {
            var writeCsv = new ReadAndWriteCsv("AlertAddressList.csv");
            writeCsv.WriteBindableDataToCsv<AlertAddressMap>(BindableAlertAddressList);
        }

        private void AlertAddressCsvImportButton_Click(object sender, EventArgs e)
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath != null)
            {
                try
                {
                    var importData = new BindingSource(
                        importAction.ReadCsv<AlertAddress>(importAction.ParseCsv<AlertAddressMap>(filePath)),
                        string.Empty);
                    foreach (var data in importData)
                    {
                        BindableAlertAddressList.Add(data);
                    }

                    MessageBox.Show("インポートが完了しました。");
                }
                catch (Exception)
                {
                    MessageBox.Show("インポートに失敗しました。");
                }
            }
        }

        private void AlertAddressCsvExportButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AlertAddressMap>(BindableAlertAddressList, "警告アドレス.csv");
        }
        #endregion

        #region AutoCcBccKeywordList setting
        public void AutoCcBccKeywordListToGrid()
        {
            var readCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");

            BindableAutoCcBccKeywordList = new BindingSource(readCsv.ReadCsv<AutoCcBccKeyword>(readCsv.ParseCsv<AutoCcBccKeywordMap>()), string.Empty);
            AutoCcBccKeywordGrid.DataSource = BindableAutoCcBccKeywordList;

            AutoCcBccKeywordGrid.Columns[0].HeaderText = @"キーワード";
            AutoCcBccKeywordGrid.Columns[1].HeaderText = @"CCまたはBCC";
            AutoCcBccKeywordGrid.Columns[2].HeaderText = @"追加アドレス";
        }

        public void SaveAutoCcBccKeywordListToCsv()
        {
            var writeCsv = new ReadAndWriteCsv("AutoCcBccKeywordList.csv");
            writeCsv.WriteBindableDataToCsv<AutoCcBccKeywordMap>(BindableAutoCcBccKeywordList);
        }

        private void AutoCcBccKeywordImportButton_Click(object sender, EventArgs e)
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath != null)
            {
                try
                {
                    var importData = new BindingSource(
                        importAction.ReadCsv<AutoCcBccKeyword>(importAction.ParseCsv<AutoCcBccKeywordMap>(filePath)),
                        string.Empty);
                    foreach (var data in importData)
                    {
                        BindableAutoCcBccKeywordList.Add(data);
                    }

                    MessageBox.Show("インポートが完了しました。");
                }
                catch (Exception)
                {
                    MessageBox.Show("インポートに失敗しました。");
                }
            }
        }

        private void AutoCcBccKeywordExportButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AutoCcBccKeywordMap>(BindableAutoCcBccKeywordList, "自動CCBCC追加キーワードリスト.csv");
        }
        #endregion

        #region AutoCcBccRecipientList setting
        public void AutoCcBccRecipientListToGrid()
        {
            var readCsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");

            BindableAutoCcBccRecipientList = new BindingSource(readCsv.ReadCsv<AutoCcBccRecipient>(readCsv.ParseCsv<AutoCcBccRecipientMap>()), string.Empty);
            AutoCcBccRecipientGrid.DataSource = BindableAutoCcBccRecipientList;

            AutoCcBccRecipientGrid.Columns[0].HeaderText = @"宛先アドレスまたはドメイン";
            AutoCcBccRecipientGrid.Columns[1].HeaderText = @"CCまたはBCC";
            AutoCcBccRecipientGrid.Columns[2].HeaderText = @"追加アドレス";
        }

        public void SaveAutoCcBccRecipientListToCsv()
        {
            var writeCsv = new ReadAndWriteCsv("AutoCcBccRecipientList.csv");
            writeCsv.WriteBindableDataToCsv<AutoCcBccRecipientMap>(BindableAutoCcBccRecipientList);
        }

        private void AutoCcBccRecipientImportCsvButton_Click(object sender, EventArgs e)
        {
            var importAction = new CsvImportAndExport();
            var filePath = importAction.ImportCsv();

            if (filePath != null)
            {
                try
                {
                    var importData = new BindingSource(
                        importAction.ReadCsv<AutoCcBccRecipient>(
                            importAction.ParseCsv<AutoCcBccRecipientMap>(filePath)), string.Empty);
                    foreach (var data in importData)
                    {
                        BindableAutoCcBccRecipientList.Add(data);
                    }

                    MessageBox.Show("インポートが完了しました。");
                }
                catch (Exception)
                {
                    MessageBox.Show("インポートに失敗しました。");
                }
            }
        }

        private void AutoCcBccRecipientExportCsvButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<AutoCcBccRecipientMap>(BindableAutoCcBccRecipientList, "自動CCBCC追加宛先リスト.csv");
        }
        #endregion

        #region Buttons.
        private void OkButton_Click(object sender, EventArgs e)
        {
            DoSaveSettings();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            //Do Nothing.
        }

        private void ApplyButton_Click(object sender, EventArgs e)
        {
            DoSaveSettings();
        }

        private void DoSaveSettings()
        {
            SaveWhitelistToCsv();
            SaveNameAndDomainsListToCsv();
            SaveAlertKeywordAndMessageListToCsv();
            SaveAlertAddressListToCsv();
            SaveAutoCcBccKeywordListToCsv();
            SaveAutoCcBccRecipientListToCsv();
        }
        #endregion
    }
}