using CsvHelper.Configuration;

namespace OutlookAddIn
{
    public class Whitelist
    {
        public string WhiteName { get; set; }
    }

    public sealed class WhitelistMap : CsvClassMap<Whitelist>
    {
        public WhitelistMap()
        {
            Map(m => m.WhiteName).Index(0);
        }
    }
}