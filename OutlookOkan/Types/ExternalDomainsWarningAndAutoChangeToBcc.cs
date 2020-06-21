using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class ExternalDomainsWarningAndAutoChangeToBcc
    {
        public int TargetToAndCcExternalDomainsNum { get; set; }
        public bool IsWarningWhenLargeNumberOfExternalDomains { get; set; }
        public bool IsProhibitedWhenLargeNumberOfExternalDomains { get; set; }
        public bool IsAutoChangeToBccWhenLargeNumberOfExternalDomains { get; set; }
    }

    public sealed class ExternalDomainsWarningAndAutoChangeToBccMap : ClassMap<ExternalDomainsWarningAndAutoChangeToBcc>
    {
        public ExternalDomainsWarningAndAutoChangeToBccMap()
        {
            Map(m => m.TargetToAndCcExternalDomainsNum).Index(0).Default(10);

            Map(m => m.IsWarningWhenLargeNumberOfExternalDomains).Index(1).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            Map(m => m.IsProhibitedWhenLargeNumberOfExternalDomains).Index(2).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            Map(m => m.IsAutoChangeToBccWhenLargeNumberOfExternalDomains).Index(3).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}