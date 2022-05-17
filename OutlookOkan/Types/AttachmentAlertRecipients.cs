using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AttachmentAlertRecipients
    {
        public string Recipient { get; set; }
        public string Message { get; set; }
    }

    public sealed class AttachmentAlertRecipientsMap : ClassMap<AttachmentAlertRecipients>
    {
        public AttachmentAlertRecipientsMap()
        {
            _ = Map(m => m.Recipient).Index(0);
            _ = Map(m => m.Message).Index(1);
        }
    }
}