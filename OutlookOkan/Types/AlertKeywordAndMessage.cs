using CsvHelper.Configuration;

namespace OutlookOkan
{
    public class AlertKeywordAndMessage
    {
        public string AlertKeyword { get; set; }
        public string Message { get; set; }
    }

    public sealed class AlertKeywordAndMessageMap : CsvClassMap<AlertKeywordAndMessage>
    {
        public AlertKeywordAndMessageMap()
        {
            Map(m => m.AlertKeyword).Index(0);
            Map(m => m.Message).Index(1);
        }
    }
}