using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class AutoCcBccRecipient
    {
        public string TargetRecipient { get; set; }
        public CcOrBcc CcOrBcc { get; set; }
        public string AutoAddAddress { get; set; }
    }

    public sealed class AutoCcBccRecipientMap : ClassMap<AutoCcBccRecipient>
    {
        public AutoCcBccRecipientMap()
        {
            Map(m => m.TargetRecipient).Index(0);
            Map(m => m.CcOrBcc).Index(1);
            Map(m => m.AutoAddAddress).Index(2);
        }
    }
}