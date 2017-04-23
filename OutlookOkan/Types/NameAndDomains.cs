using CsvHelper.Configuration;

namespace OutlookOkan
{
    public class NameAndDomains
    {
        public string Name { get; set; }
        public string Domain { get; set; }
    }

    public sealed class NameAndDomainsMap : CsvClassMap<NameAndDomains>
    {
        public NameAndDomainsMap()
        {
            Map(m => m.Name).Index(0);
            Map(m => m.Domain).Index(1);
        }
    }
}