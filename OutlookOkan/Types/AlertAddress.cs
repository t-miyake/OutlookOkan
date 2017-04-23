using CsvHelper.Configuration;

namespace OutlookOkan
{
    public class AlertAddress
    {
        public string TartgetAddress { get; set; }
    }

    public sealed class AlertAddressMap : CsvClassMap<AlertAddress>
    {
        public AlertAddressMap()
        {
            Map(m => m.TartgetAddress).Index(0);
        }
    }
}