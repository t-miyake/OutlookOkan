using System.Windows;
using OutlookOkan.ViewModels;

namespace OutlookOkan.Views
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            DataContext = new AboutWindowViewModel();
            InitializeComponent();
        }
    }
}