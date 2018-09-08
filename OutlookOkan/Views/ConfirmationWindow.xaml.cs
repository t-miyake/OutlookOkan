using OutlookOkan.Types;
using OutlookOkan.ViewModels;
using System.Windows;

namespace OutlookOkan.Views
{
    public partial class ConfirmationWindow : Window
    {
        public ConfirmationWindow(CheckList checkList)
        {
            var viewModel = new ConfirmationWindowViewModel(checkList);
            DataContext = viewModel;

            InitializeComponent();
        }

        /// <summary>
        /// DialogResultをBindしずらいので、コードビハインドで。
        /// </summary>
        private void SendButton_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        /// <summary>
        /// DialogResultをBindしずらいので、コードビハインドで。
        /// </summary>
        private void CancelButton_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        /// <summary>
        /// チェックボックスのイベント処理がやりづらかったので、コードビハインド側からViewModelのメソッドを呼び出す。
        /// </summary>
        private void ToggleButton_OnChecked(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel.ToggleSendButton();
        }

        /// <summary>
        /// チェックボックスのイベント処理がやりづらかったので、コードビハインド側からViewModelのメソッドを呼び出す。
        /// </summary>
        private void ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel.ToggleSendButton();
        }
    }
}