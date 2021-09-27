using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class ExternalDomainsWarningAndAutoChangeToBcc
    {
        public int TargetToAndCcExternalDomainsNum { get; set; } = 10;
        public bool IsWarningWhenLargeNumberOfExternalDomains { get; set; } = true;
        public bool IsProhibitedWhenLargeNumberOfExternalDomains { get; set; }
        public bool IsAutoChangeToBccWhenLargeNumberOfExternalDomains { get; set; }
    }

    public sealed class ExternalDomainsWarningAndAutoChangeToBccMap : ClassMap<ExternalDomainsWarningAndAutoChangeToBcc>
    {
        public ExternalDomainsWarningAndAutoChangeToBccMap()
        {
            _ = Map(m => m.TargetToAndCcExternalDomainsNum).Index(0).Default(10);

            _ = Map(m => m.IsWarningWhenLargeNumberOfExternalDomains).Index(1).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            _ = Map(m => m.IsProhibitedWhenLargeNumberOfExternalDomains).Index(2).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoChangeToBccWhenLargeNumberOfExternalDomains).Index(3).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}