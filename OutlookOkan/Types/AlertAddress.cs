using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AlertAddress
    {
        public string TartgetAddress { get; set; }
        public bool IsCanNotSend { get; set; }
    }

    public sealed class AlertAddressMap : ClassMap<AlertAddress>
    {
        public AlertAddressMap()
        {
            Map(m => m.TartgetAddress).Index(0);
            Map(m => m.IsCanNotSend).Index(1).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}