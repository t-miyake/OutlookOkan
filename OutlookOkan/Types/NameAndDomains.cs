using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class NameAndDomains
    {
        public string Name { get; set; }
        public string Domain { get; set; }
    }

    public sealed class NameAndDomainsMap : ClassMap<NameAndDomains>
    {
        public NameAndDomainsMap()
        {
            Map(m => m.Name).Index(0);
            Map(m => m.Domain).Index(1);
        }
    }
}