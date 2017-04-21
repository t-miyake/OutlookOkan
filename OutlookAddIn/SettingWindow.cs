using System;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class SettingWindow : Form
    {
        public SettingWindow()
        {
            InitializeComponent();

            SetNameAndDomainsListToGrid();
            WhitelistToGrid();
        }

        public BindingSource BindableWhitelist { get; set; }
        public BindingSource BindableNameAdnDomainList { get; set; }
        public BindingSource BindableAlertAndMessage { get; set; }
        public BindingSource BindableAlertAddress { get; set; }
        public BindingSource BindableAutoCcBccKeyword { get; set; }
        public BindingSource BindableAutoCcBccRecipient { get; set; }


        //TODO 共通処理のMethod化。
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
                var importData = new BindingSource(importAction.ReadCsv<NameAndDomains>(importAction.ParseCsv<NameAndDomainsMap>(filePath)), string.Empty);
                foreach (var data in importData)
                {
                    BindableNameAdnDomainList.Add(data);
                }

                MessageBox.Show("インポートが完了しました。");
            }
        }

        private void NameAndDomainsCsvExportButton_Click(object sender, EventArgs e)
        {
            var exportAction = new CsvImportAndExport();
            exportAction.CsvExport<NameAndDomainsMap>(BindableNameAdnDomainList,"名称とドメインのリスト.csv");
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
            SaveNameAndDomainsListToCsv();
            SaveWhitelistToCsv();
        }
        #endregion
    }
}