using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AlertKeywordAndMessageForSubject
    {
        public string AlertKeyword { get; set; }
        public string Message { get; set; }
        public bool IsCanNotSend { get; set; }
    }

    public sealed class AlertKeywordAndMessageForSubjectMap : ClassMap<AlertKeywordAndMessageForSubject>
    {
        public AlertKeywordAndMessageForSubjectMap()
        {
            _ = Map(m => m.AlertKeyword).Index(0);
            _ = Map(m => m.Message).Index(1);
            _ = Map(m => m.IsCanNotSend).Index(2).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}