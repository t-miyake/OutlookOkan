using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class GeneralSettings
    {
        public string Culture { get; set; }
        public bool IsDoNotShowConfirmationIfAllInWhiteList { get; set; }
        public bool IsDoNotShowConfirmationIfAllInternal { get; set; }
    }

    public sealed class GeneralSettingsMap : ClassMap<GeneralSettings>
    {
        public GeneralSettingsMap()
        {
            Map(m => m.Culture).Index(0);
            Map(m => m.IsDoNotShowConfirmationIfAllInWhiteList).Index(1);
            Map(m => m.IsDoNotShowConfirmationIfAllInternal).Index(2);
        }
    }
}