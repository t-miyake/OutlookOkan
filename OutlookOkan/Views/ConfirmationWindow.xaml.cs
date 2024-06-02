using OutlookOkan.Types;
using OutlookOkan.ViewModels;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace OutlookOkan.Views
{
    public partial class ConfirmationWindow : Window
    {
        private readonly dynamic _item;
        private readonly string _tempFilePath;

        public ConfirmationWindow(CheckList checkList, dynamic item)
        {
            DataContext = new ConfirmationWindowViewModel(checkList);

            _item = item;
            _tempFilePath = checkList.TempFilePath;

            InitializeComponent();

            //送信遅延時間を表示(設定)欄に入れる。
            DeferredDeliveryMinutesBox.Text = checkList.DeferredMinutes.ToString();
            
            //縦方向の最大サイズを制限
            MaxHeight = SystemParameters.WorkArea.Height;

            //ウィンドウサイズのロード
            if (Properties.Settings.Default.ConfirmationWindowWidth != 0)
            {
                Width = Properties.Settings.Default.ConfirmationWindowWidth;
            }

            if (Properties.Settings.Default.ConfirmationWindowHeight != 0)
            {
                Height = Properties.Settings.Default.ConfirmationWindowHeight;
            }
        }

        /// <summary>
        /// DialogResultをBindしずらいので、コードビハインドで。
        /// </summary>
        private void SendButton_OnClick(object sender, RoutedEventArgs e)
        {
            //送信時刻の設定
            _ = int.TryParse(DeferredDeliveryMinutesBox.Text, out var deferredDeliveryMinutes);

            if (deferredDeliveryMinutes != 0)
            {
                if (_item.DeferredDeliveryTime == new DateTime(4501, 1, 1, 0, 0, 0))
                {
                    //アドインの機能のみで保留時間が設定された場合
                    _item.DeferredDeliveryTime = DateTime.Now.AddMinutes(deferredDeliveryMinutes);
                }
                else
                {
                    //アドインの機能と同時に、Outlookの標準機能でも保留時間(配信タイミング)が設定された場合
                    if (DateTime.Now.AddMinutes(deferredDeliveryMinutes) > _item.DeferredDeliveryTime.AddMinutes(deferredDeliveryMinutes))
                    {
                        //[既に設定されている送信予定日時+アドインによる保留時間] が [現在日時+アドインによる保留時間] より前の日時となるため、後者を採用する。
                        _item.DeferredDeliveryTime = DateTime.Now.AddMinutes(deferredDeliveryMinutes);
                    }
                    else
                    {
                        _item.DeferredDeliveryTime = _item.DeferredDeliveryTime.AddMinutes(deferredDeliveryMinutes);
                    }
                }
            }

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
            viewModel?.ToggleSendButton();
        }

        /// <summary>
        /// チェックボックスのイベント処理がやりづらかったので、コードビハインド側からViewModelのメソッドを呼び出す。
        /// </summary>
        private void ToggleButton_OnUnchecked(object sender, RoutedEventArgs e)
        {
            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        /// <summary>
        /// 送信遅延時間の入力ボックスを数値のみ入力に制限する。
        /// </summary>
        private void DeferredDeliveryMinutesBox_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var regex = new Regex("[^0-9]+$");
            if (!regex.IsMatch(DeferredDeliveryMinutesBox.Text + e.Text)) return;

            DeferredDeliveryMinutesBox.Text = "0";
            e.Handled = true;
        }

        /// <summary>
        /// 送信遅延時間の入力ボックスへのペーストを無視する。(全角数字がペーストされる恐れがあるため)
        /// </summary>
        private void DeferredDeliveryMinutesBox_OnPreviewExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            if (e.Command == ApplicationCommands.Paste)
            {
                e.Handled = true;
            }
        }

        #region MouseUpEvent_OnHandler

        private void AlertGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            //左クリック以外は無視する。(CurrentItemがずれる場合があるため)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Alert)AlertGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            AlertGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void ToGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            //左クリック以外は無視する。(CurrentItemがずれる場合があるため)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Address)ToGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            ToGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void CcGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            //左クリック以外は無視する。(CurrentItemがずれる場合があるため)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Address)CcGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            CcGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void BccGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            //左クリック以外は無視する。(CurrentItemがずれる場合があるため)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Address)BccGrid.CurrentItem;
            currentItem.IsChecked = !currentItem.IsChecked;
            BccGrid.Items.Refresh();

            var viewModel = DataContext as ConfirmationWindowViewModel;
            viewModel?.ToggleSendButton();
        }

        private void AttachmentGridMouseUpEvent_OnHandler(object sender, MouseButtonEventArgs e)
        {
            //左クリック以外は無視する。(CurrentItemがずれる場合があるため)
            if (e.ChangedButton != MouseButton.Left) return;

            var currentItem = (Attachment)AttachmentGrid.CurrentItem;
            var cell = GetDataGridObject<DataGridCell>(AttachmentGrid, e.GetPosition(AttachmentGrid));
            if (cell is null) return;
            var columnIndex = cell.Column.DisplayIndex;

            if (columnIndex == 1 && currentItem.IsCanOpen)
            {
                var result = MessageBox.Show(Properties.Resources.OpenTheAttachedFile + " (" + currentItem.FileName + ")" + Environment.NewLine + Properties.Resources.ChangesInTheFileWillNotBeSaved, Properties.Resources.OpenTheAttachedFile, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes, MessageBoxOptions.ServiceNotification);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        var process = new ProcessStartInfo
                        {
                            UseShellExecute = true,
                            FileName = currentItem.FilePath,
                        };
                        Process.Start(process);
                    }
                    catch (Exception)
                    {
                        //Do Nothing.
                    }
                    finally
                    {
                        currentItem.IsChecked = true;
                        AttachmentGrid.Items.Refresh();
                        var viewModel = DataContext as ConfirmationWindowViewModel;
                        viewModel?.ToggleSendButton();
                    }
                }

            }
            else
            {
                if (!currentItem.IsNotMustOpenBeforeCheck) return;

                currentItem.IsChecked = !currentItem.IsChecked;
                AttachmentGrid.Items.Refresh();

                var viewModel = DataContext as ConfirmationWindowViewModel;
                viewModel?.ToggleSendButton();
            }
        }

        private T GetDataGridObject<T>(Visual dataGrid, Point point)
        {
            var result = default(T);
            var hitResultTest = VisualTreeHelper.HitTest(dataGrid, point);
            if (hitResultTest == null) return result;
            var visualHit = hitResultTest.VisualHit;
            while (visualHit != null)
            {
                if (visualHit is T)
                {
                    result = (T)(object)visualHit;
                    break;
                }
                visualHit = VisualTreeHelper.GetParent(visualHit);
            }
            return result;
        }

        #endregion

        private void ConfirmationWindow_OnClosing(object sender, CancelEventArgs e)
        {
            // ウインドウサイズを保存
            Properties.Settings.Default.ConfirmationWindowWidth = Width;
            Properties.Settings.Default.ConfirmationWindowHeight = Height;
            Properties.Settings.Default.Save();
            
            if (string.IsNullOrEmpty(_tempFilePath)) return;

            try
            {
                File.Delete(_tempFilePath);
            }
            catch (Exception)
            {
                // Do Nothing.
            }
        }
    }
}