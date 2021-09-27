using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AutoCcBccRecipient
    {
        public string TargetRecipient { get; set; }
        public CcOrBcc CcOrBcc { get; set; }
        public string AutoAddAddress { get; set; }
    }

    public sealed class AutoCcBccRecipientMap : ClassMap<AutoCcBccRecipient>
    {
        public AutoCcBccRecipientMap()
        {
            _ = Map(m => m.TargetRecipient).Index(0);
            _ = Map(m => m.CcOrBcc).Index(1);
            _ = Map(m => m.AutoAddAddress).Index(2);
        }
    }
}