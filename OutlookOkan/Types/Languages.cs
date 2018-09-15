using System.Collections.Generic;

namespace OutlookOkan.Types
{
    public class Languages
    {
        public List<LanguageCodeAndName> Language = new List<LanguageCodeAndName>();

        public Languages()
        {
            Language.Add(new LanguageCodeAndName { LanguageNumber = 0, LanguageName = "日本語", LanguageCode = "ja-JP" });
            Language.Add(new LanguageCodeAndName { LanguageNumber = 1, LanguageName = "English", LanguageCode = "en-US" });
        }
    }

    public class LanguageCodeAndName
    {
        public int LanguageNumber { get; set; }
        public string LanguageCode { get; set; }
        public string LanguageName { get; set; }
    }
}