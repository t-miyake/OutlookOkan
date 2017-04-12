using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemSend +=  new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        public void Application_ItemSend(object Item, ref bool Cancel)
        {

            Outlook.MailItem mail = Item as Outlook.MailItem;

            var ConfirmWindow = new ConfirmWindow(mail);
            var DialogResult = ConfirmWindow.ShowDialog();

            ConfirmWindow.Dispose();

            if(DialogResult == DialogResult.OK)
            {
                //メールを送信。
            }else if (DialogResult == DialogResult.Cancel)
            {
                Cancel = true;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //注: Outlook はこのイベントを発行しなくなりました。Outlook が
            //    シャットダウンする際に実行が必要なコードがある場合は、http://go.microsoft.com/fwlink/?LinkId=506785 を参照してください。
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

}