using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Win32;
using OutlookOkan.ViewModels;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace OutlookOkan.Views
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            DataContext = new SettingsWindowViewModel();

            InitializeComponent();
        }

        #region Validations

        private void DataGrid_WhiteList_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var inputText = ((TextBox)e.EditingElement).Text;
            if (string.IsNullOrEmpty(inputText))
            {
                _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
            else
            {
                //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                if (!inputText.Equals("@")) return;

                _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
        }

        private void DataGrid_NameAndDomains_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AlertKeywordAndMessage_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AlertKeywordAndMessageForSubject_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AlertAddresses_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            switch (e.Column.DisplayIndex)
            {
                case 0:
                    var inputText = ((TextBox)e.EditingElement).Text;
                    if (string.IsNullOrEmpty(inputText))
                    {
                        _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                        e.Cancel = true;
                    }
                    return;

                case 1:
                    return;

                default:
                    return;
            }
        }

        private void DataGrid_AutoCcBccKeyword_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AutoCcBccRecipient_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AutoCcBccAttachedFile_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_DeferredDeliveryMinutes_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            switch (e.Column.DisplayIndex)
            {
                case 0:
                    if (!string.IsNullOrEmpty(((TextBox)e.EditingElement).Text) && ((TextBox)e.EditingElement).Text.Contains("@")) return;

                    _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                    e.Cancel = true;
                    return;

                case 1:
                    var regex = new Regex("[^0-9]+$");
                    if (!regex.IsMatch(((TextBox)e.EditingElement).Text)) return;

                    _ = MessageBox.Show(Properties.Resources.InputDeferredDeliveryTime, Properties.Resources.AppName, MessageBoxButton.OK);
                    e.Cancel = true;
                    return;

                default:
                    return;
            }
        }

        private void DataGrid_InternalDomainList_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var inputText = ((TextBox)e.EditingElement).Text;
            if (string.IsNullOrEmpty(inputText) || (!inputText.StartsWith("@") && !inputText.StartsWith(".")))
            {
                _ = MessageBox.Show(Properties.Resources.InputDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
            else
            {
                //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                if (!inputText.Equals("@") && !inputText.Equals(".")) return;

                _ = MessageBox.Show(Properties.Resources.InputDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
        }

        private void DataGrid_RecipientsAndAttachmentsName_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            switch (e.Column.DisplayIndex)
            {
                case 0:
                    var inputText = ((TextBox)e.EditingElement).Text;
                    if (string.IsNullOrEmpty(inputText) || !inputText.Contains("@"))
                    {
                        _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                        e.Cancel = true;
                    }
                    else
                    {
                        //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                        if (!inputText.Equals("@")) return;

                        _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                        e.Cancel = true;
                    }
                    return;

                case 1:
                    return;

                default:
                    return;
            }
        }

        private void DataGrid_AttachmentProhibitedRecipients_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var inputText = ((TextBox)e.EditingElement).Text;
            if (string.IsNullOrEmpty(inputText) || !inputText.Contains("@"))
            {
                _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
            else
            {
                //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                if (!inputText.Equals("@")) return;

                _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
        }

        private void ExternalDomainsNumBox_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var regex = new Regex("[^0-9]+$");
            if (!regex.IsMatch(ExternalDomainsNumBox.Text + e.Text)) return;

            e.Handled = true;
        }

        private void ExternalDomainsNumBox_OnPreviewExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Command == ApplicationCommands.Paste)
            {
                e.Handled = true;
            }
        }

        private void DataGrid_AttachmentAlertRecipients_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            switch (e.Column.DisplayIndex)
            {
                case 0:
                    var inputText = ((TextBox)e.EditingElement).Text;
                    if (string.IsNullOrEmpty(inputText) || !inputText.Contains("@"))
                    {
                        _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                        e.Cancel = true;
                    }
                    //else
                    //{
                    //    //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                    //    if (!inputText.Equals("@")) return;

                    //    _ = MessageBox.Show(Properties.Resources.InputMailaddressOrDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                    //    e.Cancel = true;
                    //}
                    return;

                case 1:
                    return;

                default:
                    return;
            }
        }

        #endregion

        #region Buttons

        //エクスポート
        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Config files (*.ocfg)|*.ocfg",
                Title = Properties.Resources.ExportAllSettings,
                FileName = "OutlookAddinConfig.ocfg"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                var sourceDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
                var targetZipFilePath = saveFileDialog.FileName;
                const string password = "cWepiJ3kkc2k";

                using (var zipOutputStream = new ZipOutputStream(File.Create(targetZipFilePath)))
                {
                    zipOutputStream.SetLevel(1);
                    zipOutputStream.Password = password;

                    var buffer = new byte[4096];

                    foreach (var filePath in Directory.GetFiles(sourceDirectory))
                    {
                        var entry = new ZipEntry(Path.GetFileName(filePath))
                        {
                            DateTime = DateTime.Now
                        };
                        zipOutputStream.PutNextEntry(entry);

                        using (var fs = File.OpenRead(filePath))
                        {
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                zipOutputStream.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }

                    zipOutputStream.Finish();
                    zipOutputStream.Close();
                }
                _ = MessageBox.Show(Properties.Resources.CompletedExport);
            }
        }

        //インポート
        private void ImportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Config files (*.ocfg)|*.ocfg",
                Title = Properties.Resources.ImportAllSettings
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    // 暗号化ZIPファイルを展開し、特定のディレクトリ内のファイルを置き換える
                    var sourceZipFilePath = openFileDialog.FileName;
                    var targetDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Noraneko\\OutlookOkan\\");
                    const string password = "cWepiJ3kkc2k";

                    using (var zipInputStream = new ZipInputStream(File.OpenRead(sourceZipFilePath)))
                    {
                        zipInputStream.Password = password; // パスワード設定  
                        ZipEntry entry;

                        while ((entry = zipInputStream.GetNextEntry()) != null)
                        {
                            var targetFilePath = Path.Combine(targetDirectory, entry.Name);

                            using (var fileStream = File.Create(targetFilePath))
                            {
                                var buffer = new byte[2048];

                                while (true)
                                {
                                    var size = zipInputStream.Read(buffer, 0, buffer.Length);
                                    if (size > 0)
                                    {
                                        fileStream.Write(buffer, 0, size);
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }

                        zipInputStream.Close();
                    }

                    _ = MessageBox.Show(Properties.Resources.CompletedImport);
                    Close();

                }
                catch (Exception ex)
                {
                    _ = MessageBox.Show(Properties.Resources.ImportErrorOfAllSettings + ex.Message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as SettingsWindowViewModel;
            _ = viewModel?.SaveSettings();

            DialogResult = true;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as SettingsWindowViewModel;
            _ = viewModel?.SaveSettings();

            _ = MessageBox.Show(Properties.Resources.SaveSettings, Properties.Resources.AppName, MessageBoxButton.OK);
        }

        #endregion
    }
}