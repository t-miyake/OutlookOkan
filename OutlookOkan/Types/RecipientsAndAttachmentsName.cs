using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class RecipientsAndAttachmentsName
    {
        public string AttachmentsName { get; set; }
        public string Recipient { get; set; }
    }

    public sealed class RecipientsAndAttachmentsNameMap : ClassMap<RecipientsAndAttachmentsName>
    {
        public RecipientsAndAttachmentsNameMap()
        {
            _ = Map(m => m.AttachmentsName).Index(0);
            _ = Map(m => m.Recipient).Index(1);
        }
    }
}