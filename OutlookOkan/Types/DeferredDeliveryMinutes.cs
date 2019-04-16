using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class DeferredDeliveryMinutes
    {
        public string TartgetAddress { get; set; }
        public int DeferredMinutes { get; set; }
    }

    public sealed class DeferredDeliveryMinutesMap : ClassMap<DeferredDeliveryMinutes>
    {
        public DeferredDeliveryMinutesMap()
        {
            Map(m => m.TartgetAddress).Index(0);
            Map(m => m.DeferredMinutes).Index(1).Default(0);
        }
    }
}