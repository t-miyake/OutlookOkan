using System;
using System.Collections;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace SetupCustomAction
{
    /// <summary>
    /// インストーラ用のカスタムアクション
    /// </summary>
    [System.ComponentModel.RunInstaller(true)]
    public sealed class CustomAction : Installer
    {
        /// <summary>
        /// インストール時のカスタムアクション
        /// </summary>
        /// <param name="savedState">savedState</param>
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            //msiexec /i "OkanSetup.msi" SILENT=TRUE ALLUSERS=1 /quiet /norestart
            //ALLUSERS=1 で、すべてのユーザを対象にインストール
            if (Context.Parameters["silent"] == "TRUE") return;

            var outlookProcess = Process.GetProcessesByName("OUTLOOK");
            if (outlookProcess.Length <= 0) return;

            _ = MessageBox.Show("Outlookが起動しています。Outlookを終了してからインストールしてください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK, MessageBoxOptions.ServiceNotification);
            throw new InstallException();
        }

        /// <summary>
        /// アンインストール時のカスタムアクション
        /// </summary>
        /// <param name="savedState">savedState</param>
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            try
            {
                //msiexec /x "OkanSetup.msi" DELCONF=TRUE /quiet /norestart
                if (Context.Parameters["delconf"] == "TRUE")
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

                    return;
                }

                //msiexec /x "OkanSetup.msi" SILENT=TRUE /quiet /norestart
                if (Context.Parameters["silent"] == "TRUE") return;

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
            catch (Exception)
            {
                //Do Nothing.
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