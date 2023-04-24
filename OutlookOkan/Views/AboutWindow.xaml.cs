using OutlookOkan.ViewModels;
using System.Windows;

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