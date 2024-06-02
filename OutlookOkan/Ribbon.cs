using OutlookOkan.Handlers;
using OutlookOkan.Helpers;
using OutlookOkan.Views;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using MessageBox = System.Windows.MessageBox;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public void ShowHelp(Office.IRibbonControl control)
        {
            _ = Process.Start("https://github.com/t-miyake/OutlookOkan/wiki/Manual");
        }

        public void ShowSettings(Office.IRibbonControl control)
        {
            var settingsWindow = new SettingsWindow();
            var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
            var outlookHandle = new NativeMethods(activeWindow).Handle;
            _ = new WindowInteropHelper(settingsWindow) { Owner = outlookHandle };

            _ = settingsWindow.ShowDialog();
        }

        public void ShowAbout(Office.IRibbonControl control)
        {
            var aboutWindow = new AboutWindow();
            var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
            var outlookHandle = new NativeMethods(activeWindow).Handle;
            _ = new WindowInteropHelper(aboutWindow) { Owner = outlookHandle };

            _ = aboutWindow.ShowDialog();
        }

        public void VerifyEmailHeader(Office.IRibbonControl control)
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (!(explorer.Selection[1] is Outlook.MailItem mailItem)) return;

                var headers = mailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E").ToString();
                if (string.IsNullOrEmpty(headers)) return;

                var analysisResults = MailHeaderHandler.ValidateEmailHeader(headers);
                var selfGeneratedDmarcResult = MailHeaderHandler.DetermineDmarcResult(analysisResults["SPF"], analysisResults["SPF Alignment"], analysisResults["DKIM"], analysisResults["DKIM Alignment"]);

                var message = "";
                foreach (KeyValuePair<string, string> entry in analysisResults)
                {
                    message += ($"{entry.Key}: {entry.Value}") + Environment.NewLine;
                }
                message += ($"DMARC By OKan: {selfGeneratedDmarcResult}") + Environment.NewLine;

                if (analysisResults["Internal"] == "TRUE")
                {
                    message = "Internal: TRUE";
                    _ = MessageBox.Show(message, Properties.Resources.AppName, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    if (analysisResults["DMARC"] != "PASS" && analysisResults["DMARC"] != "BESTGUESSPASS" && selfGeneratedDmarcResult == "FAIL")
                    {
                        _ = MessageBox.Show(Properties.Resources.SpoofingRiskWaring + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        _ = MessageBox.Show(message, Properties.Resources.AppName, MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception)
            {
                //Do Nothing.
            }
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
                case "VerifyEmailHeader":
                    result = Properties.Resources.VerifyEmailHeader;
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
