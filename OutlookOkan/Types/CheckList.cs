using System.Collections.Generic;

namespace OutlookOkan.Types
{
    public sealed class CheckList
    {
        public List<Alert> Alerts { get; set; } = new List<Alert>();
        public List<Address> ToAddresses { get; set; } = new List<Address>();
        public List<Address> CcAddresses { get; set; } = new List<Address>();
        public List<Address> BccAddresses { get; set; } = new List<Address>();
        public List<Attachment> Attachments { get; set; } = new List<Attachment>();
        public string Sender { get; set; }
        public string SenderDomain { get; set; }
        public int RecipientExternalDomainNumAll { get; set; }
        public string Subject { get; set; }
        public string MailType { get; set; }
        public string MailBody { get; set; }
        public string MailHtmlBody { get; set; }
        public bool IsCanNotSendMail { get; set; }
        public string CanNotSendMailMessage { get; set; }
        public int DeferredMinutes { get; set; }
    }

    public sealed class Alert
    {
        public string AlertMessage { get; set; }
        public bool IsImportant { get; set; }
        public bool IsWhite { get; set; }
        public bool IsChecked { get; set; }
    }

    public sealed class Attachment
    {
        public string FileName { get; set; }
        public string FileType { get; set; }
        public string FileSize { get; set; }
        public bool IsTooBig { get; set; }
        public bool IsDangerous { get; set; }
        public bool IsEncrypted { get; set; }
        public bool IsChecked { get; set; }
    }

    public sealed class Address
    {
        public string MailAddress { get; set; }
        public bool IsExternal { get; set; }
        public bool IsWhite { get; set; }
        public bool IsSkip { get; set; }
        public bool IsChecked { get; set; }
    }
}