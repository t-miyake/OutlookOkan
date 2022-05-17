using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AlertAddress
    {
        public string TargetAddress { get; set; }
        public bool IsCanNotSend { get; set; }
        public string Message { get; set; }
    }

    public sealed class AlertAddressMap : ClassMap<AlertAddress>
    {
        public AlertAddressMap()
        {
            _ = Map(m => m.TargetAddress).Index(0);
            _ = Map(m => m.IsCanNotSend).Index(1).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            _ = Map(m => m.Message).Index(2);
        }
    }
}