using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class Whitelist
    {
        public string WhiteName { get; set; }
        public bool IsSkipConfirmation { get; set; }
    }

    public sealed class WhitelistMap : ClassMap<Whitelist>
    {
        public WhitelistMap()
        {
            _ = Map(m => m.WhiteName).Index(0);
            _ = Map(m => m.IsSkipConfirmation).Index(1).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}