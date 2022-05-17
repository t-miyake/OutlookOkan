using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AutoAddMessage
    {
        public bool IsAddToStart { get; set; }
        public bool IsAddToEnd { get; set; }
        public string MessageOfAddToStart { get; set; }
        public string MessageOfAddToEnd { get; set; }
    }

    public sealed class AutoAddMessageMap : ClassMap<AutoAddMessage>
    {
        public AutoAddMessageMap()
        {
            _ = Map(m => m.IsAddToStart).Index(0).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            _ = Map(m => m.IsAddToEnd).Index(1).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            _ = Map(m => m.MessageOfAddToStart).Index(2);
            _ = Map(m => m.MessageOfAddToEnd).Index(3);
        }
    }
}