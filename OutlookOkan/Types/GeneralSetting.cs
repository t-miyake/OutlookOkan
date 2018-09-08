using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public class GeneralSetting
    {
        public bool IsDoNotConfirmationIfAllRecipientsAreSameDomain { get; set; }
        public bool IsDoDoNotConfirmationIfAllWhite { get; set; }
        public bool IsAutoCheckIfAllRecipientsAreSameDomain { get; set; }
        public string LanguageCode { get; set; }
        public bool IsShowConfirmationToMultipleDomain { get; set; }
    }

    public sealed class GeneralSettingMap : ClassMap<GeneralSetting>
    {
        public GeneralSettingMap()
        {
            Map(m => m.IsDoNotConfirmationIfAllRecipientsAreSameDomain).Index(0).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            Map(m => m.IsDoDoNotConfirmationIfAllWhite).Index(1).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            Map(m => m.IsAutoCheckIfAllRecipientsAreSameDomain).Index(2).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            Map(m => m.LanguageCode).Index(3);
            Map(m => m.IsShowConfirmationToMultipleDomain).Index(4).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}