using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public Ribbon(){}

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("OutlookAddIn.Ribbon.xml");
        }

        #endregion

        #region リボンのコールバック
        //ここでコールバック メソッドを作成します。コールバック メソッドの追加について詳しくは https://go.microsoft.com/fwlink/?LinkID=271226 をご覧ください

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            this._ribbon = ribbonUi;
        }

        public void ShowSettings(Office.IRibbonControl control)
        {
            var settingWindow = new SettingWindow();
            var temp = settingWindow.ShowDialog();
            settingWindow.Dispose();
        }

        public void ShowVersion(Office.IRibbonControl control)
        {
            var aboutBox = new AboutBox();
            var temp = aboutBox.ShowDialog();
            aboutBox.Dispose();
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}