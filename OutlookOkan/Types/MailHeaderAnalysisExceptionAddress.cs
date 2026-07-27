using CsvHelper.Configuration;

namespace OutlookOkan.Types
{
    /// <summary>
    /// メールヘッダ解析(なりすまし/SPF/DKIM/DMARC)警告の例外とする送信者メールアドレス。
    /// ここに完全一致するアドレスからの受信メールは、ヘッダ解析による警告を表示しない。
    /// </summary>
    public class MailHeaderAnalysisExceptionAddress
    {
        public string TargetAddress { get; set; }
    }

    public sealed class MailHeaderAnalysisExceptionAddressMap : ClassMap<MailHeaderAnalysisExceptionAddress>
    {
        public MailHeaderAnalysisExceptionAddressMap()
        {
            _ = Map(m => m.TargetAddress).Index(0);
        }
    }
}
