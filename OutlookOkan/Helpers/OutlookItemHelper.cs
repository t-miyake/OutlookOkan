using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Helpers
{
    internal static class OutlookItemHelper
    {
        //MAPIプロパティタグ。
        private const string PrTransportMessageHeaders = @"http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        private const string PrSenderSmtpAddress = @"http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
        private const string PrSentRepresentingSmtpAddress = @"http://schemas.microsoft.com/mapi/proptag/0x5D02001F";
        private const string PrSmtpAddress = @"http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        /// <summary>
        /// メールのインターネットヘッダ(PR_TRANSPORT_MESSAGE_HEADERS)を取得する。
        /// </summary>
        /// <param name="mail">対象のメールアイテム</param>
        /// <returns>ヘッダ文字列。取得できない場合は null</returns>
        internal static string TryGetTransportHeaders(Outlook.MailItem mail)
        {
            if (mail is null) return null;

            try
            {
                return mail.PropertyAccessor.GetProperty(PrTransportMessageHeaders) as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 受信メールの送信者のSMTPアドレスを取得する。
        /// </summary>
        /// <param name="mail">対象のメールアイテム</param>
        /// <returns>送信者のSMTPアドレス。取得できない場合は null</returns>
        internal static string TryGetSenderSmtpAddress(Outlook.MailItem mail)
        {
            if (mail is null) return null;

            //1) SMTPアカウント由来のメールは SenderEmailAddress がそのままSMTPアドレス
            try
            {
                if (string.Equals(mail.SenderEmailType, "SMTP", StringComparison.OrdinalIgnoreCase) && IsLikelyEmailAddress(mail.SenderEmailAddress))
                {
                    return mail.SenderEmailAddress;
                }
            }
            catch (Exception)
            {
                //次へ
            }

            //2) PropertyAccessor からSMTPアドレスを取得 (Exchange環境等)
            foreach (var propertyTag in new[] { PrSenderSmtpAddress, PrSentRepresentingSmtpAddress, PrSmtpAddress })
            {
                try
                {
                    if (mail.PropertyAccessor.GetProperty(propertyTag) is string address && IsLikelyEmailAddress(address))
                    {
                        return address;
                    }
                }
                catch (Exception)
                {
                    //次へ
                }
            }

            //3) Exchangeユーザとして解決 (Exchange環境のみ)
            try
            {
                var smtpAddress = mail.Sender?.GetExchangeUser()?.PrimarySmtpAddress;
                if (IsLikelyEmailAddress(smtpAddress)) return smtpAddress;
            }
            catch (Exception)
            {
                //次へ
            }

            //4) 最後のフォールバック (Exchange DN等、SMTP形式でない可能性あり)
            try
            {
                if (IsLikelyEmailAddress(mail.SenderEmailAddress)) return mail.SenderEmailAddress;
            }
            catch (Exception)
            {
                //取得不能
            }

            return null;
        }

        /// <summary>
        /// メールアドレスらしい文字列か(最低限 @ を含むか)を判定する。
        /// </summary>
        private static bool IsLikelyEmailAddress(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && value.Contains("@");
        }
    }
}
