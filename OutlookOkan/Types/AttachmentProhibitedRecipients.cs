using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AttachmentProhibitedRecipients
    {
        public string Recipient { get; set; }
    }

    public sealed class AttachmentProhibitedRecipientsMap : ClassMap<AttachmentProhibitedRecipients>
    {
        public AttachmentProhibitedRecipientsMap()
        {
            _ = Map(m => m.Recipient).Index(0);
        }
    }
}