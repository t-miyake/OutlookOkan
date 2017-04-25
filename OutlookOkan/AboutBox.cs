using System.Reflection;
using System.Windows.Forms;

namespace OutlookOkan
{
    internal partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();
            Text = $"{AssemblyTitle} のバージョン情報";
            labelProductName.Text = AssemblyProduct;
            labelVersion.Text = $"バージョン {AssemblyVersion}";
            labelCopyright.Text = AssemblyCopyright;
            labelCompanyName.Text = AssemblyCompany;
            textBoxDescription.Text = "Third-Party Software Usage and Licenses\r\n\r\nCsvHelper\r\nCopyright © 2009-2017 Josh Close and Contributors\r\nDual licensed under Microsoft Public License and Apache License Version 2.0\r\nhttps://github.com/JoshClose/CsvHelper/blob/master/LICENSE.txt";
        }

        #region アセンブリ属性アクセサー

        public string AssemblyTitle
        {
            get
            {
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    var titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion => Assembly.GetExecutingAssembly().GetName().Version.ToString();

        public string AssemblyProduct
        {
            get
            {
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                var attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                return attributes.Length == 0 ? "" : ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion
    }
}