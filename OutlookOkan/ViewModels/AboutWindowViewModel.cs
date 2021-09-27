using OutlookOkan.Models;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;

namespace OutlookOkan.ViewModels
{
    internal sealed class AboutWindowViewModel : ViewModelBase
    {
        private readonly CheckNewVersion _checkNewVersion = new CheckNewVersion();

        internal AboutWindowViewModel()
        {
            CheckNewVersionButtonCommand = new RelayCommand(CheckNewVersion);
        }

        public ICommand CheckNewVersionButtonCommand { get; }

        private void CheckNewVersion()
        {
            if (_checkNewVersion.IsCanDownloadNewVersion())
            {
                var result = MessageBox.Show(Properties.Resources.CanGetNewVersion, Properties.Resources.AppName, MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    _ = Process.Start("https://github.com/t-miyake/OutlookOkan/releases");
                }
            }
            else
            {
                _ = MessageBox.Show(Properties.Resources.YouHaveLatest, Properties.Resources.AppName, MessageBoxButton.OK);
            }
        }
    }
}