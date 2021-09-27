using OutlookOkan.ViewModels;
using System;
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
            try
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
            catch (Exception)
            {
                //Do Nothing.
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
            if (string.IsNullOrEmpty(inputText) || !inputText.Contains("@"))
            {
                _ = MessageBox.Show(Properties.Resources.InputDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
            else
            {
                //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                if (!inputText.Equals("@")) return;

                _ = MessageBox.Show(Properties.Resources.InputDomain, Properties.Resources.AppName, MessageBoxButton.OK);
                e.Cancel = true;
            }
        }
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

        #endregion

        #region Buttons

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