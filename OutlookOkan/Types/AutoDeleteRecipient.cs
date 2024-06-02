using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AutoDeleteRecipient
    {
        public string Recipient { get; set; }
    }

    public sealed class AutoDeleteRecipientMap : ClassMap<AutoDeleteRecipient>
    {
        public AutoDeleteRecipientMap()
        {
            _ = Map(m => m.Recipient).Index(0);
        }
    }
}
