using OutlookOkan.ViewModels;
using System;
using System.Windows;
using System.Windows.Controls;

namespace OutlookOkan.Views
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            var viewModel = new SettingsWindowViewModel();
            DataContext = viewModel;

            InitializeComponent();
        }

        #region Validations

        /// <summary>
        /// WhiteListへの入力バリデーション
        /// </summary>
        private void DataGrid_WhiteList_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            var inputText = ((TextBox)e.EditingElement).Text;
            if (string.IsNullOrEmpty(inputText)||!inputText.Contains("@"))
            {
                MessageBox.Show(Properties.Resources.InputMailaddressOrDomain);
                e.Cancel = true;
            }
            else
            {
                //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                if (!inputText.Equals("@")) return;
                MessageBox.Show(Properties.Resources.InputMailaddressOrDomain);
                e.Cancel = true;
            }
        }

        private void DataGrid_NameAndDomains_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AlertKeywordAndMessage_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
        }

        private void DataGrid_AlertAddresses_OnCellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                var inputText = ((TextBox) e.EditingElement).Text;
                if (string.IsNullOrEmpty(inputText) || !inputText.Contains("@"))
                {
                    MessageBox.Show(Properties.Resources.InputMailaddressOrDomain);
                    e.Cancel = true;
                }
                else
                {
                    //@のみで登録すると全てのメールアドレスが対象になるため、それを禁止。
                    if (!inputText.Equals("@")) return;
                    MessageBox.Show(Properties.Resources.InputMailaddressOrDomain);
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

        #endregion

        #region Buttons

        private void OkButton_OnClick(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as SettingsWindowViewModel;
            viewModel.SaveSettings();

            DialogResult = true;
        }

        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void ApplyButton_OnClick(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as SettingsWindowViewModel;
            viewModel.SaveSettings();

            MessageBox.Show(Properties.Resources.SaveSettings);
        }

        #endregion
    }
}