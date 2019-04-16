using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AlertKeywordAndMessage
    {
        public string AlertKeyword { get; set; }
        public string Message { get; set; }
        public bool IsCanNotSend { get; set; }
    }

    public sealed class AlertKeywordAndMessageMap : ClassMap<AlertKeywordAndMessage>
    {
        public AlertKeywordAndMessageMap()
        {
            Map(m => m.AlertKeyword).Index(0);
            Map(m => m.Message).Index(1);
            Map(m => m.IsCanNotSend).Index(2).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}