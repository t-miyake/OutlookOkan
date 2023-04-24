using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class AlertKeywordOfSubjectWhenOpeningMail
    {
        public string AlertKeyword { get; set; }
        public string Message { get; set; }
    }

    public sealed class AlertKeywordOfSubjectWhenOpeningMailMap : ClassMap<AlertKeywordOfSubjectWhenOpeningMail>
    {
        public AlertKeywordOfSubjectWhenOpeningMailMap()
        {
            _ = Map(m => m.AlertKeyword).Index(0);
            _ = Map(m => m.Message).Index(1);
        }
    }
}