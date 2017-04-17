using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public Ribbon()
        {
        }

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

        public void ShowVersion(Office.IRibbonControl control)
        {
            MessageBox.Show(@"誤送信防止アドイン"+ Environment.NewLine + @"Version 0.2" + Environment.NewLine + @"(株)のらねこ", @"バージョン情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
