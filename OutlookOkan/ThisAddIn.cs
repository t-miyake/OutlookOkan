using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookOkan
{
    public partial class ThisAddIn
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
              return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend += Application_ItemSend;
        }

        private static void Application_ItemSend(object item, ref bool cancel)
        {
            var confirmationWindow = new ConfirmationWindow(item as Outlook._MailItem);
            var dialogResult = confirmationWindow.ShowDialog();

            confirmationWindow.Dispose();

            if(dialogResult == DialogResult.OK)
            {
                //メールを送信。
            }else
            {
                cancel = true;
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
        }

        #endregion
    }

}