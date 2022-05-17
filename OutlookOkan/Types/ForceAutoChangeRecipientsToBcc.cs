using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class ForceAutoChangeRecipientsToBcc
    {
        public bool IsForceAutoChangeRecipientsToBcc { get; set; }
        public string ToRecipient { get; set; }
        public bool IsIncludeInternalDomain { get; set; }
    }

    public sealed class ForceAutoChangeRecipientsToBccMap : ClassMap<ForceAutoChangeRecipientsToBcc>
    {
        public ForceAutoChangeRecipientsToBccMap()
        {
            _ = Map(m => m.IsForceAutoChangeRecipientsToBcc).Index(0).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            _ = Map(m => m.ToRecipient).Index(1);
            _ = Map(m => m.IsIncludeInternalDomain).Index(2).TypeConverterOption.BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}