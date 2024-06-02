using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class SecurityForReceivedMail
    {
        public bool IsEnableSecurityForReceivedMail { get; set; }
        public bool IsEnableAlertKeywordOfSubjectWhenOpeningMailsData { get; set; }
        public bool IsEnableMailHeaderAnalysis { get; set; }
        public bool IsShowWarningWhenSpfFails { get; set; }
        public bool IsShowWarningWhenDkimFails { get; set; }
        public bool IsEnableWarningFeatureWhenOpeningAttachments { get; set; }
        public bool IsWarnBeforeOpeningAttachments { get; set; }
        public bool IsWarnBeforeOpeningEncryptedZip { get; set; }
        public bool IsWarnLinkFileInTheZip { get; set; }
        public bool IsWarnOneFileInTheZip { get; set; }
        public bool IsWarnOfficeFileWithMacroInTheZip { get; set; }
        public bool IsWarnBeforeOpeningAttachmentsThatContainMacros { get; set; }
        public bool IsShowWarningWhenSpoofingRisk { get; set; }
        public bool IsShowWarningWhenDmarcNotImplemented { get; set; }
    }

    public sealed class SecurityForReceivedMailMap : ClassMap<SecurityForReceivedMail>
    {
        public SecurityForReceivedMailMap()
        {
            _ = Map(m => m.IsEnableSecurityForReceivedMail).Index(0).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsEnableAlertKeywordOfSubjectWhenOpeningMailsData).Index(1).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsEnableMailHeaderAnalysis).Index(2).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsShowWarningWhenSpfFails).Index(3).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsShowWarningWhenDkimFails).Index(4).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsEnableWarningFeatureWhenOpeningAttachments).Index(5).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarnBeforeOpeningAttachments).Index(6).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarnBeforeOpeningEncryptedZip).Index(7).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarnLinkFileInTheZip).Index(8).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarnOneFileInTheZip).Index(9).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarnOfficeFileWithMacroInTheZip).Index(10).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarnBeforeOpeningAttachmentsThatContainMacros).Index(11).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsShowWarningWhenSpoofingRisk).Index(12).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsShowWarningWhenDmarcNotImplemented).Index(13).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}