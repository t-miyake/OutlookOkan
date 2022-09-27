using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class KeywordAndRecipients
    {
        public string Keyword { get; set; }
        public string Recipient { get; set; }
    }

    public sealed class KeywordAndRecipientsMap : ClassMap<KeywordAndRecipients>
    {
        public KeywordAndRecipientsMap()
        {
            _ = Map(m => m.Keyword).Index(0);
            _ = Map(m => m.Recipient).Index(1);
        }
    }
}