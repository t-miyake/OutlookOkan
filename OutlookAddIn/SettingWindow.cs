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
        }

        public BindingSource BindableNameAdnDomainList { get; set; }

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


        private void OkButton_Click(object sender, EventArgs e)
        {
            SaveNameAndDomainsListToCsv();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            //Do Nothing.
        }

        private void ApplyButton_Click(object sender, EventArgs e)
        {
            SaveNameAndDomainsListToCsv();
        }
    }
}