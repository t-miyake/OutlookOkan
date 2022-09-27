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
        public bool EnableForgottenToAttachAlert { get; set; } = true;
        public bool EnableGetContactGroupMembers { get; set; }
        public bool EnableGetExchangeDistributionListMembers { get; set; }
        public bool ContactGroupMembersAreWhite { get; set; } = true;
        public bool ExchangeDistributionListMembersAreWhite { get; set; } = true;
        public bool IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles { get; set; }
        public bool IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain { get; set; }
        public bool IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain { get; set; }
        public bool IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain { get; set; }
        public bool IsEnableRecipientsAreSortedByDomain { get; set; }
        public bool IsAutoAddSenderToBcc { get; set; }
        public bool IsAutoCheckRegisteredInContacts { get; set; }
        public bool IsAutoCheckRegisteredInContactsAndMemberOfContactLists { get; set; }
        public bool IsCheckNameAndDomainsFromRecipients { get; set; }
        public bool IsWarningIfRecipientsIsNotRegistered { get; set; }
        public bool IsProhibitsSendingMailIfRecipientsIsNotRegistered { get; set; }
        public bool IsShowConfirmationAtSendMeetingRequest { get; set; }
        public bool IsAutoAddSenderToCc { get; set; }
        public bool IsCheckNameAndDomainsIncludeSubject { get; set; }
        public bool IsCheckNameAndDomainsFromSubject { get; set; }
        public bool IsShowConfirmationAtSendTaskRequest { get; set; }
        public bool IsAutoCheckAttachments { get; set; }
        public bool IsCheckKeywordAndRecipientsIncludeSubject { get; set; }
    }

    public sealed class GeneralSettingMap : ClassMap<GeneralSetting>
    {
        public GeneralSettingMap()
        {
            _ = Map(m => m.IsDoNotConfirmationIfAllRecipientsAreSameDomain).Index(0).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsDoDoNotConfirmationIfAllWhite).Index(1).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoCheckIfAllRecipientsAreSameDomain).Index(2).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.LanguageCode).Index(3);

            _ = Map(m => m.IsShowConfirmationToMultipleDomain).Index(4).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.EnableForgottenToAttachAlert).Index(5).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            _ = Map(m => m.EnableGetContactGroupMembers).Index(6).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.EnableGetExchangeDistributionListMembers).Index(7).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.ContactGroupMembersAreWhite).Index(8).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            _ = Map(m => m.ExchangeDistributionListMembersAreWhite).Index(9).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(true);

            _ = Map(m => m.IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles).Index(10).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain).Index(11).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain).Index(12).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain).Index(13).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsEnableRecipientsAreSortedByDomain).Index(14).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoAddSenderToBcc).Index(15).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoCheckRegisteredInContacts).Index(16).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoCheckRegisteredInContactsAndMemberOfContactLists).Index(17).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsCheckNameAndDomainsFromRecipients).Index(18).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarningIfRecipientsIsNotRegistered).Index(19).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsProhibitsSendingMailIfRecipientsIsNotRegistered).Index(20).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsShowConfirmationAtSendMeetingRequest).Index(21).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoAddSenderToCc).Index(22).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsCheckNameAndDomainsIncludeSubject).Index(23).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsCheckNameAndDomainsFromSubject).Index(24).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsShowConfirmationAtSendTaskRequest).Index(25).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAutoCheckAttachments).Index(26).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsCheckKeywordAndRecipientsIncludeSubject).Index(27).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}