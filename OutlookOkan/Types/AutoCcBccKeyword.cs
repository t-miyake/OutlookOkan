using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AutoCcBccKeyword
    {
        public string Keyword { get; set; }
        public CcOrBcc CcOrBcc { get; set; }
        public string AutoAddAddress { get; set; }
    }

    public sealed class AutoCcBccKeywordMap : ClassMap<AutoCcBccKeyword>
    {
        public AutoCcBccKeywordMap()
        {
            _ = Map(m => m.Keyword).Index(0);
            _ = Map(m => m.CcOrBcc).Index(1);
            _ = Map(m => m.AutoAddAddress).Index(2);
        }
    }
}