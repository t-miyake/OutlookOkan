using System;
using System.Collections;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace SetupCustomAction
{
    [System.ComponentModel.RunInstaller(true)]
    public sealed class CustomAction : Installer
    {
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            var outlookProcess = Process.GetProcessesByName("OUTLOOK");
            if (outlookProcess.Length > 0)
            {
                throw new InstallException();
            }
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            var result = MessageBox.Show("設定を削除しますか？", "設定削除の確認", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.ServiceNotification);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
                    Directory.Delete(directoryPath, true);
                }
                catch (Exception)
                {
                    //Do Nothing.
                }
            }
        }

        public override void Commit(IDictionary savedState)
        {
        }

        public override void Rollback(IDictionary savedState)
        {
        }
    }
}
