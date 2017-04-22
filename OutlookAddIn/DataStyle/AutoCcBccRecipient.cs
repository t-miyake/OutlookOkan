using CsvHelper.Configuration;

namespace OutlookAddIn
{
    public class AutoCcBccRecipient
    {
        public string TargetRecipient { get; set; }
        public CcOrBcc CcOrBcc { get; set; }
        public string AutoAddAddress { get; set; }
    }
    public sealed class AutoCcBccRecipientMap : CsvClassMap<AutoCcBccRecipient>
    {
        public AutoCcBccRecipientMap()
        {
            Map(m => m.TargetRecipient).Index(0);
            Map(m => m.CcOrBcc).Index(1);
            Map(m => m.AutoAddAddress).Index(2);
        }
    }
}