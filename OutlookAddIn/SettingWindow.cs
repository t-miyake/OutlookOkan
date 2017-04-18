using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CsvHelper;

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
        readonly string SettingPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookAddin\\");
        private const string SettingFile = "NameAndDomains.csv";

        //TODO CSV読み込み/書き込みを他でも使う可能性があるため、それぞの機能はメソッドとして分離させる。
        public void SetNameAndDomainsListToGrid()
        {
            if (!Directory.Exists(SettingPath))
                Directory.CreateDirectory(SettingPath);

            if (!File.Exists(SettingPath + SettingFile))
                File.Create(SettingPath + SettingFile).Close();

            using (var csvParser =
                new CsvParser(new StreamReader(SettingPath + SettingFile, Encoding.GetEncoding("Shift_JIS"))))
            {
                csvParser.Configuration.HasHeaderRecord = false;
                csvParser.Configuration.RegisterClassMap<NameAndDomainsMap>();

                BindableNameAdnDomainList = new BindingSource(new CsvReader(csvParser).GetRecords<NameAndDomains>().ToList(), string.Empty);

                NameAndDomainsGrid.DataSource = BindableNameAdnDomainList;

                NameAndDomainsGrid.Columns[0].HeaderText = @"名称";
                NameAndDomainsGrid.Columns[1].HeaderText = @"ドメイン (@から)";
            }
        }

        public void SaveNameAndDomainsListToCsv()
        {
            using (var csvWriter = 
                new CsvWriter(new StreamWriter(SettingPath + SettingFile, false, Encoding.GetEncoding("Shift_JIS"))))
            {
                csvWriter.Configuration.HasHeaderRecord = false;
                csvWriter.Configuration.RegisterClassMap<NameAndDomainsMap>();

                csvWriter.WriteRecords(BindableNameAdnDomainList);
            }
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