using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class GeneralSetting
    {
        public bool IsDoNotConfirmationIfAllRecipientsAreSameDomain { get; set; }
        public bool IsDoDoNotConfirmationIfAllWhite { get; set; }
        public bool IsAutoCheckIfAllRecipientsAreSameDomain { get; set; }
        public string LanguageCode { get; set; }
        public bool IsShowConfirmationToMultipleDomain { get; set; }
        public bool EnableForgottenToAttachAlert { get; set; }
        public bool EnableGetContactGroupMembers { get; set; }
        public bool EnableGetExchangeDistributionListMembers { get; set; }
        public bool ContactGroupMembersAreWhite { get; set; }
        public bool ExchangeDistributionListMembersAreWhite { get; set; }
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

            Map(m => m.EnableForgottenToAttachAlert).Index(5).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            Map(m => m.EnableGetContactGroupMembers).Index(6).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            Map(m => m.EnableGetExchangeDistributionListMembers).Index(7).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            Map(m => m.ContactGroupMembersAreWhite).Index(8).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            Map(m => m.ExchangeDistributionListMembersAreWhite).Index(9).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

        }
    }
}