using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class Whitelist
    {
        public string WhiteName { get; set; }
    }

    public sealed class WhitelistMap : ClassMap<Whitelist>
    {
        public WhitelistMap()
        {
            Map(m => m.WhiteName).Index(0);
        }
    }
}