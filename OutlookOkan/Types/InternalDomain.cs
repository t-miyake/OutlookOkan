using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class InternalDomain
    {
        public string Domain { get; set; }
    }

    public sealed class InternalDomainMap : ClassMap<InternalDomain>
    {
        public InternalDomainMap()
        {
            _ = Map(m => m.Domain).Index(0);
        }
    }
}