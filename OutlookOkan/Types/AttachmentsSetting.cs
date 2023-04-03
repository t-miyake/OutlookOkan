using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    public sealed class AttachmentsSetting
    {
        public bool IsWarningWhenEncryptedZipIsAttached { get; set; }
        public bool IsProhibitedWhenEncryptedZipIsAttached { get; set; }
        public bool IsEnableAllAttachedFilesAreDetectEncryptedZip { get; set; }
        public bool IsAttachmentsProhibited { get; set; }
        public bool IsWarningWhenAttachedRealFile { get; set; }
        public bool IsEnableOpenAttachedFiles { get; set; }
        public string TargetAttachmentFileExtensionOfOpen { get; set; } = ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx";
        public bool IsMustOpenBeforeCheckTheAttachedFiles { get; set; }
        public bool IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain { get; set; }
    }

    public sealed class AttachmentsSettingMap : ClassMap<AttachmentsSetting>
    {
        public AttachmentsSettingMap()
        {
            _ = Map(m => m.IsWarningWhenEncryptedZipIsAttached).Index(0).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsProhibitedWhenEncryptedZipIsAttached).Index(1).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsEnableAllAttachedFilesAreDetectEncryptedZip).Index(2).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsAttachmentsProhibited).Index(3).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsWarningWhenAttachedRealFile).Index(4).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.IsEnableOpenAttachedFiles).Index(5).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);

            _ = Map(m => m.TargetAttachmentFileExtensionOfOpen).Index(6).Default(".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx");

            _ = Map(m => m.IsMustOpenBeforeCheckTheAttachedFiles).Index(7).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
            
            _ = Map(m => m.IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain).Index(8).TypeConverterOption
                .BooleanValues(true, true, "Yes", "Y").TypeConverterOption.BooleanValues(false, true, "No", "N").Default(false);
        }
    }
}