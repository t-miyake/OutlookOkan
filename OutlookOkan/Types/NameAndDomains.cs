using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class NameAndDomains
    {
        public string Name { get; set; }
        public string Domain { get; set; }
    }

    public sealed class NameAndDomainsMap : ClassMap<NameAndDomains>
    {
        public NameAndDomainsMap()
        {
            _ = Map(m => m.Name).Index(0);
            _ = Map(m => m.Domain).Index(1);
        }
    }
}