using OutlookOkan.Helpers;
using OutlookOkan.Views;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Office = Microsoft.Office.Core;

namespace OutlookOkan
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public void ShowHelp(Office.IRibbonControl control)
        {
            Process.Start("https://github.com/t-miyake/OutlookOkan/wiki/Manual");
        }

        public void ShowSettings(Office.IRibbonControl control)
        {
            var settingsWindow = new SettingsWindow();
            var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
            var outlookHandle = new NativeMethods(activeWindow).Handle;
            _ = new WindowInteropHelper(settingsWindow) { Owner = outlookHandle };

            settingsWindow.ShowDialog();
        }

        public void ShowAbout(Office.IRibbonControl control)
        {
            var aboutWindow = new AboutWindow();
            var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
            var outlookHandle = new NativeMethods(activeWindow).Handle;
            _ = new WindowInteropHelper(aboutWindow) { Owner = outlookHandle };

            aboutWindow.ShowDialog();
        }

        /// <summary>
        /// リボンの多言語化処理
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public string GetLabel(Office.IRibbonControl control)
        {
            string result = null;
            switch (control.Id)
            {
                case "Settings":
                    result = Properties.Resources.Settings;
                    break;
                case "About":
                    result = Properties.Resources.About;
                    break;
                case "Help":
                    result = Properties.Resources.Help;
                    break;
                case "MyAddinGroup":
                    result = Properties.Resources.AppName;
                    break;
                default:
                    break;
            }
            return result;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonId)
        {
            return ribbonId == "Microsoft.Outlook.Explorer" ? GetResourceText("OutlookOkan.Ribbon.xml") : null;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            foreach (var t in resourceNames.Where(t => string.Compare(resourceName, t, StringComparison.OrdinalIgnoreCase) == 0))
            {
                using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(t) ?? throw new InvalidOperationException()))
                {
                    return resourceReader.ReadToEnd();
                }
            }
            return null;
        }

        #endregion
    }
}
