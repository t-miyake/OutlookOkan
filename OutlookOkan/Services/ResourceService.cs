using OutlookOkan.Properties;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace OutlookOkan.Services
{
    class ResourceService : INotifyPropertyChanged
    {
        // Singleton instance.
        public static ResourceService Instance { get; } = new ResourceService();
        private ResourceService(){}

        public Resources Resources { get; } = new Resources();

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void RaisePropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void ChangeCulture(string name)
        {
            Resources.Culture = CultureInfo.GetCultureInfo(name);
            RaisePropertyChanged("Resources");
        }
    }
}