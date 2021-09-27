using System.Collections.Generic;

namespace OutlookOkan.Types
{
    public sealed class Languages
    {
        public List<LanguageCodeAndName> Language = new List<LanguageCodeAndName>();

        public Languages()
        {
            Language.Add(new LanguageCodeAndName { LanguageNumber = 0, LanguageName = "日本語 [Japanese]", LanguageCode = "ja-JP" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 1, LanguageName = "English", LanguageCode = "en-US" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 2, LanguageName = "简体中文 [Chinese (Simplified, China)] Beta", LanguageCode = "zh-CN" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 3, LanguageName = "繁體中文 [Chinese (Traditional, Taiwan)] Beta", LanguageCode = "zh-TW" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 4, LanguageName = "Español [Spanish] Beta", LanguageCode = "es-ES" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 5, LanguageName = "हिन्दी [Hindi] Beta", LanguageCode = "hi-IN" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 6, LanguageName = "Русский язык [Russian] Beta", LanguageCode = "ru-RU" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 7, LanguageName = "Deutsch [German] Beta", LanguageCode = "de-DE" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 8, LanguageName = "한국어 [Korean] Beta", LanguageCode = "ko-KR" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 9, LanguageName = "ไทย [Thai] Beta", LanguageCode = "th-TH" });
        }
    }

    public sealed class LanguageCodeAndName
    {
        public int LanguageNumber { get; set; }
        public string LanguageCode { get; set; }
        public string LanguageName { get; set; }
    }
}