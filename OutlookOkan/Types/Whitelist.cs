using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class Whitelist
    {
        public string WhiteName { get; set; }
        public bool IsSkipConfirmation { get; set; }
    }

    public sealed class WhitelistMap : ClassMap<Whitelist>
    {
        public WhitelistMap()
        {
            Map(m => m.WhiteName).Index(0);
            Map(m => m.IsSkipConfirmation).Index(1).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}