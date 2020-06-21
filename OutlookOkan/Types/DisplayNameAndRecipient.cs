using System.Collections.Generic;

namespace OutlookOkan.Types
{
    public sealed class DisplayNameAndRecipient
    {
        public Dictionary<string, string> All { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> To { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> Cc { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> Bcc { get; set; } = new Dictionary<string, string>();
        public List<MailItemsRecipientAndMailAddress> MailRecipientsIndex { get; set; } = new List<MailItemsRecipientAndMailAddress>();
    }
}