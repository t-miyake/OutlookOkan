using Microsoft.VisualStudio.TestTools.UnitTesting;
using OutlookOkan.Models;
using OutlookOkan.Properties;
using OutlookOkan.Types;
using OutlookOkanTest.Types;
using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkanTest
{
    [TestClass]
    public class UnitTest
    {
        #region _GenerateCheckList

        #region GetSenderAndSenderDomain

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetSenderAndSenderDomain")]
        public void 送信元のアドレスとドメインの取得_通常送信_送信元がExchange()
        {
            //NOTE 通常は存在しえないケース。
            //NOTE テストを実行する環境で正常に情報取得可能なExchangeアカウントでのみテストが成功するため、環境によって適切なアドレスに変更すること。
            const string exchangeEmailAddress = "miyake@noraneko.co.jp";

            var testMailItem = new TestMailItem { SenderEmailType = "EX", SenderEmailAddress = exchangeEmailAddress };
            var checkList = new CheckList();

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem, checkList };
            var result = (CheckList)privateObject.Invoke("GetSenderAndSenderDomain", args);

            Console.WriteLine(@"Sender email address：" + result.Sender);
            Console.WriteLine(@"Sender domain：" + result.SenderDomain);

            Assert.AreEqual(exchangeEmailAddress, result.Sender);
            Assert.AreEqual(exchangeEmailAddress.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal)), result.SenderDomain);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetSenderAndSenderDomain")]
        public void 送信元のアドレスとドメインの取得_通常送信_送信元がExchangeのCN()
        {
            //NOTE テストを実行する環境で正常に情報取得可能なExchangeアカウントでのみテストが成功するため、環境によって適切なアドレスに変更すること。
            const string exchangeEmailAddress = "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=97915151601144a1a18a149c879b05ad-mailbox1";
            const string plainEmailAddress = "miyake@noraneko.co.jp";

            var testMailItem = new TestMailItem { SenderEmailType = "EX", SenderEmailAddress = exchangeEmailAddress };
            var checkList = new CheckList();

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem, checkList };
            var result = (CheckList)privateObject.Invoke("GetSenderAndSenderDomain", args);

            Console.WriteLine(@"Sender email address：" + result.Sender);
            Console.WriteLine(@"Sender domain：" + result.SenderDomain);

            Assert.AreEqual(plainEmailAddress, result.Sender);
            Assert.AreEqual(plainEmailAddress.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal)), result.SenderDomain);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetSenderAndSenderDomain")]
        public void 送信元のアドレスとドメインの取得_通常送信_送信元がExchange以外()
        {
            const string senderEmailAddress = "test@sample.com";

            var testMailItem = new TestMailItem { SenderEmailAddress = senderEmailAddress };
            var checkList = new CheckList();

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem, checkList };
            var result = (CheckList)privateObject.Invoke("GetSenderAndSenderDomain", args);

            Console.WriteLine(@"Sender email address：" + result.Sender);
            Console.WriteLine(@"Sender domain：" + result.SenderDomain);

            Assert.AreEqual(senderEmailAddress, result.Sender);
            Assert.AreEqual(senderEmailAddress.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal)), result.SenderDomain);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetSenderAndSenderDomain")]
        public void 送信元のアドレスとドメインの取得_代理送信_xxx()
        {
            //TODO
        }

        #endregion

        #region MakeEmbeddedAttachmentsList

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("MakeEmbeddedAttachmentsList")]
        public void メール本文に埋め込まれた画像などのファイル名のリストを取得_テキスト形式()
        {
            var testMailItem = new TestMailItem { BodyFormat = Outlook.OlBodyFormat.olFormatPlain };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem };
            var result = (List<string>)privateObject.Invoke("MakeEmbeddedAttachmentsList", args);

            Assert.IsNull(result);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("MakeEmbeddedAttachmentsList")]
        public void メール本文に埋め込まれた画像などのファイル名のリストを取得_リッチテキスト形式()
        {
            var testMailItem = new TestMailItem { BodyFormat = Outlook.OlBodyFormat.olFormatPlain };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem };
            var result = (List<string>)privateObject.Invoke("MakeEmbeddedAttachmentsList", args);

            Assert.IsNull(result);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("MakeEmbeddedAttachmentsList")]
        public void メール本文に埋め込まれた画像などのファイル名のリストを取得_HTML形式に埋め込みあり()
        {
            var testMailItem = new TestMailItem
            {
                BodyFormat = Outlook.OlBodyFormat.olFormatHTML,
                HTMLBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"><!--[if !mso]><style>v\\:* {behavior:url(#default#VML);}\r\no\\:* {behavior:url(#default#VML);}\r\nw\\:* {behavior:url(#default#VML);}\r\n.shape {behavior:url(#default#VML);}\r\n</style><![endif]--><style><!--\r\n/* Font Definitions */\r\n@font-face\r\n\t{font-family:\"Cambria Math\";\r\n\tpanose-1:2 4 5 3 5 4 6 3 2 4;}\r\n@font-face\r\n\t{font-family:\"Yu Gothic\";\r\n\tpanose-1:2 11 4 0 0 0 0 0 0 0;}\r\n@font-face\r\n\t{font-family:Calibri;\r\n\tpanose-1:2 15 5 2 2 2 4 3 2 4;}\r\n@font-face\r\n\t{font-family:\"\\@Yu Gothic\";\r\n\tpanose-1:2 11 4 0 0 0 0 0 0 0;}\r\n/* Style Definitions */\r\np.MsoNormal, li.MsoNormal, div.MsoNormal\r\n\t{margin:0mm;\r\n\tmargin-bottom:.0001pt;\r\n\ttext-align:justify;\r\n\ttext-justify:inter-ideograph;\r\n\tfont-size:10.5pt;\r\n\tfont-family:\"Calibri\",sans-serif;}\r\nspan.EmailStyle17\r\n\t{mso-style-type:personal-compose;\r\n\tfont-family:\"Calibri\",sans-serif;\r\n\tcolor:windowtext;}\r\n.MsoChpDefault\r\n\t{mso-style-type:export-only;\r\n\tfont-family:\"Calibri\",sans-serif;}\r\n.MsoPapDefault\r\n\t{mso-style-type:export-only;\r\n\ttext-align:justify;\r\n\ttext-justify:inter-ideograph;}\r\n/* Page Definitions */\r\n@page WordSection1\r\n\t{size:612.0pt 792.0pt;\r\n\tmargin:72.0pt 72.0pt 72.0pt 72.0pt;}\r\ndiv.WordSection1\r\n\t{page:WordSection1;}\r\n--></style><!--[if gte mso 9]><xml>\r\n<o:shapedefaults v:ext=\"edit\" spidmax=\"1026\">\r\n<v:textbox inset=\"5.85pt,.7pt,5.85pt,.7pt\" />\r\n</o:shapedefaults></xml><![endif]--><!--[if gte mso 9]><xml>\r\n<o:shapelayout v:ext=\"edit\">\r\n<o:idmap v:ext=\"edit\" data=\"1\" />\r\n</o:shapelayout></xml><![endif]--></head><body lang=JA link=\"#0563C1\" vlink=\"#954F72\" style='text-justify-trim:punctuation'><div class=WordSection1><p class=MsoNormal><span lang=EN-US style='font-size:11.0pt'><img width=512 height=512 style='width:5.3333in;height:5.3333in' id=\"Picture_x0020_1\" src=\"cid:image001.png@01D63C48.6A912FA0\" alt=\"A close up of a logo&#10;&#10;Description automatically generated\"></span><span lang=EN-US style='font-size:11.0pt'><o:p></o:p></span></p><p class=MsoNormal><span lang=EN-US style='font-size:11.0pt'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN-US style='font-size:11.0pt'><img width=2113 height=1309 style='width:22.0069in;height:13.6388in' id=\"Picture_x0020_3\" src=\"cid:image002.jpg@01D63C48.A7DCB6E0\" alt=\"A close up of text on a white background&#10;&#10;Description automatically generated\"></span><span lang=EN-US style='font-size:11.0pt'><o:p></o:p></span></p></div></body></html>"
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem };
            var result = (List<string>)privateObject.Invoke("MakeEmbeddedAttachmentsList", args);

            Console.WriteLine(@"Embedded attachments: " + result.Count);
            foreach (var attachmentName in result)
            {
                Console.WriteLine(@"AttachmentName: " + attachmentName);
            }

            Assert.AreEqual(2, result.Count);
            CollectionAssert.AreEquivalent(result, new List<string> { "image001.png", "image002.jpg" });
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("MakeEmbeddedAttachmentsList")]
        public void メール本文に埋め込まれた画像などのファイル名のリストを取得_HTML形式に埋め込みなし()
        {
            var testMailItem = new TestMailItem
            {
                BodyFormat = Outlook.OlBodyFormat.olFormatHTML,
                HTMLBody = "<html xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w=\"urn:schemas-microsoft-com:office:word\" xmlns:m=\"http://schemas.microsoft.com/office/2004/12/omml\" xmlns=\"http://www.w3.org/TR/REC-html40\"><head><meta name=Generator content=\"Microsoft Word 15 (filtered medium)\"><style><!--\r\n/* Font Definitions */\r\n@font-face\r\n\t{font-family:\"Cambria Math\";\r\n\tpanose-1:2 4 5 3 5 4 6 3 2 4;}\r\n@font-face\r\n\t{font-family:\"Yu Gothic\";\r\n\tpanose-1:2 11 4 0 0 0 0 0 0 0;}\r\n@font-face\r\n\t{font-family:Calibri;\r\n\tpanose-1:2 15 5 2 2 2 4 3 2 4;}\r\n@font-face\r\n\t{font-family:\"\\@Yu Gothic\";\r\n\tpanose-1:2 11 4 0 0 0 0 0 0 0;}\r\n/* Style Definitions */\r\np.MsoNormal, li.MsoNormal, div.MsoNormal\r\n\t{margin:0mm;\r\n\tmargin-bottom:.0001pt;\r\n\ttext-align:justify;\r\n\ttext-justify:inter-ideograph;\r\n\tfont-size:10.5pt;\r\n\tfont-family:\"Calibri\",sans-serif;}\r\nspan.EmailStyle17\r\n\t{mso-style-type:personal-compose;\r\n\tfont-family:\"Calibri\",sans-serif;\r\n\tcolor:windowtext;}\r\n.MsoChpDefault\r\n\t{mso-style-type:export-only;\r\n\tfont-family:\"Calibri\",sans-serif;}\r\n.MsoPapDefault\r\n\t{mso-style-type:export-only;\r\n\ttext-align:justify;\r\n\ttext-justify:inter-ideograph;}\r\n/* Page Definitions */\r\n@page WordSection1\r\n\t{size:612.0pt 792.0pt;\r\n\tmargin:72.0pt 72.0pt 72.0pt 72.0pt;}\r\ndiv.WordSection1\r\n\t{page:WordSection1;}\r\n--></style><!--[if gte mso 9]><xml>\r\n<o:shapedefaults v:ext=\"edit\" spidmax=\"1026\">\r\n<v:textbox inset=\"5.85pt,.7pt,5.85pt,.7pt\" />\r\n</o:shapedefaults></xml><![endif]--><!--[if gte mso 9]><xml>\r\n<o:shapelayout v:ext=\"edit\">\r\n<o:idmap v:ext=\"edit\" data=\"1\" />\r\n</o:shapelayout></xml><![endif]--></head><body lang=JA link=\"#0563C1\" vlink=\"#954F72\" style='text-justify-trim:punctuation'><div class=WordSection1><p class=MsoNormal><span lang=EN-US style='font-size:11.0pt'>test</span><span lang=EN-US><o:p></o:p></span></p></div></body></html>"
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testMailItem };
            var result = (List<string>)privateObject.Invoke("MakeEmbeddedAttachmentsList", args);

            Assert.IsNull(result);
        }

        #endregion

        #region GetMailBodyFormat

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetMailBodyFormat")]
        public void メール本文のタイプから表記を取得_テキスト形式()
        {
            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { Outlook.OlBodyFormat.olFormatPlain };
            var result = (string)privateObject.Invoke("GetMailBodyFormat", args);

            Assert.AreEqual(Resources.Text, result);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetMailBodyFormat")]
        public void メール本文のタイプから表記を取得_null()
        {
            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { null };
            var result = (string)privateObject.Invoke("GetMailBodyFormat", args);

            Assert.AreEqual(Resources.Unknown, result);
        }

        #endregion

        #region GetMailBody

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetMailBody")]
        public void テキスト形式のメール本文を取得_HTML形式()
        {
            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { Outlook.OlBodyFormat.olFormatHTML, "aa\r\n\r\nbb" };
            var result = (string)privateObject.Invoke("GetMailBody", args);

            Assert.AreEqual("aa\r\nbb", result);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetMailBody")]
        public void テキスト形式のメール本文を取得_Text形式()
        {
            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { Outlook.OlBodyFormat.olFormatPlain, "aa\r\n\r\nbb" };
            var result = (string)privateObject.Invoke("GetMailBody", args);

            Assert.AreEqual("aa\r\n\r\nbb", result);
        }

        #endregion

        #region GetAttachmentsInformation

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetAttachmentsInformation")]
        public void GetAttachmentsInformation()
        {
            //TODO
        }

        #endregion

        #region MakeDisplayNameAndRecipient

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("MakeDisplayNameAndRecipient")]
        public void MakeDisplayNameAndRecipient()
        {
            //TODO
        }

        #endregion

        #region GetNameAndRecipient

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetNameAndRecipient")]
        public void 表示名とメールアドレスを表示用に変換_Nameがメールアドレス()
        {
            const string recipient = "test@sample.com";
            const string nameAndMailAddress = "test@sample.com (test@sample.com)";
            var testRecipient = new TestRecipient { Name = recipient };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testRecipient };
            var result = (List<NameAndRecipient>)privateObject.Invoke("GetNameAndRecipient", args);

            Assert.AreEqual(result[0].MailAddress, recipient);
            Assert.AreEqual(result[0].NameAndMailAddress, nameAndMailAddress);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetNameAndRecipient")]
        public void 表示名とメールアドレスを表示用に変換_Nameが名称()
        {
            const string recipient = "test";
            var nameAndMailAddress = $"test ({Resources.FailedToGetInformation})";
            var testRecipient = new TestRecipient { Name = recipient };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testRecipient };
            var result = (List<NameAndRecipient>)privateObject.Invoke("GetNameAndRecipient", args);

            Assert.AreEqual(result[0].MailAddress, nameAndMailAddress);
            Assert.AreEqual(result[0].NameAndMailAddress, nameAndMailAddress);
        }

        #endregion

        #region CheckForgotAttach

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CheckForgotAttach")]
        public void CheckForgotAttach()
        {
            //TODO
        }

        #endregion

        #region CheckKeyword

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CheckKeyword")]
        public void 警告キーワードのチェック_該当キーワードが1件あり()
        {
            const string alertKeyword = "TEST";
            const string message = "TESTが含まれます。";
            var testAlertKeywordAndMessages = new List<AlertKeywordAndMessage>
            {
                new AlertKeywordAndMessage {AlertKeyword = alertKeyword, Message = message, IsCanNotSend = false}
            };
            var testCheckList = new CheckList { MailBody = "TEST あいうえお" };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, testAlertKeywordAndMessages };
            var result = (CheckList)privateObject.Invoke("CheckKeyword", args);

            Assert.AreEqual(result.Alerts[0].AlertMessage, message);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CheckKeyword")]
        public void 警告キーワードのチェック_該当キーワードが3件あり()
        {
            const string alertKeyword0 = "TEST";
            const string message0 = "TESTが含まれます。";
            const string alertKeyword1 = "あいうえお";
            const string message1 = "あいうえおが含まれます。";
            const string alertKeyword2 = "さしすせそ";
            const string message2 = "さしすせそが含まれます。";

            var testAlertKeywordAndMessages = new List<AlertKeywordAndMessage>
            {
                new AlertKeywordAndMessage {AlertKeyword = alertKeyword0, Message = message0, IsCanNotSend = false},
                new AlertKeywordAndMessage {AlertKeyword = alertKeyword1, Message = message1, IsCanNotSend = false},
                new AlertKeywordAndMessage {AlertKeyword = alertKeyword2, Message = message2, IsCanNotSend = true},
                new AlertKeywordAndMessage {AlertKeyword = "ほげほげ", Message = "ほげほげ", IsCanNotSend = true},
            };
            var testCheckList = new CheckList { MailBody = "TEST あいうえお さしすせそ なにぬねの ほげ" };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, testAlertKeywordAndMessages };
            var result = (CheckList)privateObject.Invoke("CheckKeyword", args);

            Assert.AreEqual(result.Alerts[0].AlertMessage, message0);
            Assert.AreEqual(result.Alerts[1].AlertMessage, message1);
            Assert.AreEqual(result.Alerts[2].AlertMessage, message2);
        }

        #endregion

        #region AutoAddCcAndBcc

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("AutoAddCcAndBcc")]
        public void AutoAddCcAndBcc()
        {
            //TODO
        }

        #endregion

        #region GetRecipient

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetRecipient")]
        public void 宛先を表示リストに格納_警告アドレスあり()
        {
            const string to1 = "三宅 (miyake@noraneko.co.jp)";
            const string to2 = "たろう (taro@sample.com)";
            const string cc1 = "(株)のらねこ (info@noraneko.co.jp)";

            var internalDomainList = new List<InternalDomain>();

            var testCheckList = new CheckList { SenderDomain = "@noraneko.co.jp" };
            var displayNameAndRecipient = new DisplayNameAndRecipient
            {
                To =
                {
                    ["miyake@noraneko.co.jp"] = to1,
                    ["taro@sample.com"] = to2
                },
                Cc = { ["info@noraneko.co.jp"] = cc1 }
            };

            var alertAddressList = new List<AlertAddress>
            {
                new AlertAddress {TargetAddress = "@sample.com", IsCanNotSend = false}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, displayNameAndRecipient, alertAddressList, internalDomainList };
            var result = (CheckList)privateObject.Invoke("GetRecipient", args);

            Assert.AreEqual(result.ToAddresses[0].MailAddress, to1);
            Assert.AreEqual(result.ToAddresses[1].MailAddress, to2);
            Assert.AreEqual(result.CcAddresses[0].MailAddress, cc1);

            Assert.AreEqual(result.Alerts[0].AlertMessage, Resources.IsAlertAddressToAlert + $"[{to2}]");
            Assert.IsFalse(testCheckList.IsCanNotSendMail);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetRecipient")]
        public void 宛先を表示リストに格納_送信禁止アドレスあり()
        {
            const string to1 = "三宅 (miyake@noraneko.co.jp)";
            const string to2 = "たろう (taro@sample.com)";
            const string cc1 = "(株)のらねこ (info@noraneko.co.jp)";

            var internalDomainList = new List<InternalDomain>();

            var testCheckList = new CheckList { SenderDomain = "@noraneko.co.jp" };
            var displayNameAndRecipient = new DisplayNameAndRecipient
            {
                To =
                {
                    ["miyake@noraneko.co.jp"] = to1,
                    ["taro@sample.com"] = to2
                },
                Cc = { ["info@noraneko.co.jp"] = cc1 }
            };

            var alertAddressList = new List<AlertAddress>
            {
                new AlertAddress {TargetAddress = "@sample.com", IsCanNotSend = true}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, displayNameAndRecipient, alertAddressList, internalDomainList };
            var result = (CheckList)privateObject.Invoke("GetRecipient", args);

            Assert.AreEqual(result.ToAddresses[0].MailAddress, to1);
            Assert.AreEqual(result.ToAddresses[1].MailAddress, to2);
            Assert.AreEqual(result.CcAddresses[0].MailAddress, cc1);

            Assert.AreEqual(result.Alerts[0].AlertMessage, Resources.IsAlertAddressToAlert + $"[{to2}]");
            Assert.IsTrue(testCheckList.IsCanNotSendMail);
            Assert.AreEqual(testCheckList.CanNotSendMailMessage, Resources.SendingForbidAddress + $"[{to2}]");
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetRecipient")]
        public void 宛先を表示リストに格納_内部ドメイン設定あり()
        {
            const string to1 = "三宅 (miyake@noraneko.co.jp)";
            const string to2 = "たろう (taro@sample.com)";
            const string cc1 = "(株)のらねこ (info@noraneko.co.jp)";
            const string bcc1 = "次郎 (jiro@sample.com)";

            var internalDomainList = new List<InternalDomain> { new InternalDomain { Domain = "@sample.com" } };

            var testCheckList = new CheckList { SenderDomain = "@noraneko.co.jp" };
            var displayNameAndRecipient = new DisplayNameAndRecipient
            {
                To =
                {
                    ["miyake@noraneko.co.jp"] = to1,
                    ["taro@sample.com"] = to2
                },
                Cc = { ["info@noraneko.co.jp"] = cc1 },
                Bcc = { ["jiro@sample2.com"] = bcc1 }
            };

            var alertAddressList = new List<AlertAddress>();

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, displayNameAndRecipient, alertAddressList, internalDomainList };
            var result = (CheckList)privateObject.Invoke("GetRecipient", args);

            Assert.AreEqual(result.ToAddresses[0].MailAddress, to1);
            Assert.AreEqual(result.ToAddresses[1].MailAddress, to2);
            Assert.AreEqual(result.CcAddresses[0].MailAddress, cc1);
            Assert.AreEqual(result.BccAddresses[0].MailAddress, bcc1);

            Assert.IsFalse(result.ToAddresses[0].IsExternal && result.ToAddresses[1].IsExternal && result.CcAddresses[0].IsExternal);
            Assert.IsTrue(result.BccAddresses[0].IsExternal);
        }

        #endregion

        #region CheckMailBodyAndRecipient

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CheckMailBodyAndRecipient")]
        public void 名称と宛先の紐づけ確認_警告あり()
        {
            var testCheckList = new CheckList { SenderDomain = "@noraneko.co.jp", MailBody = "ほげほげ株式会社" };
            var displayNameAndRecipient = new DisplayNameAndRecipient
            {
                All = { ["taro@sample.com"] = "たろう (taro@sample.com)", ["info@noraneko.co.jp"] = "(株)のらねこ (info@noraneko.co.jp)" }
            };
            var nameAndDomainsList = new List<NameAndDomains>
            {
                new NameAndDomains {Name = "ほげほげ株式会社", Domain = "@sample.hogehoge"}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, displayNameAndRecipient, nameAndDomainsList };
            var result = (CheckList)privateObject.Invoke("CheckMailBodyAndRecipient", args);

            Assert.AreEqual(result.Alerts[0].AlertMessage, "たろう (taro@sample.com)" + " : " + Resources.IsAlertAddressMaybeIrrelevant);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CheckMailBodyAndRecipient")]
        public void 名称と宛先の紐づけ確認_警告なし()
        {
            var testCheckList = new CheckList { SenderDomain = "@noraneko.co.jp", MailBody = "ほげほげ株式会社" };
            var displayNameAndRecipient = new DisplayNameAndRecipient
            {
                All = { ["taro@sample.com"] = "たろう (taro@sample.com)", ["info@noraneko.co.jp"] = "(株)のらねこ (info@noraneko.co.jp)" }
            };
            var nameAndDomainsList = new List<NameAndDomains>
            {
                new NameAndDomains {Name = "ふがふが株式会社", Domain = "@sample.fugafuga"}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testCheckList, displayNameAndRecipient, nameAndDomainsList };
            var result = (CheckList)privateObject.Invoke("CheckMailBodyAndRecipient", args);

            Assert.AreEqual(result.Alerts.Count, 0);
        }

        #endregion

        #region CountRecipientExternalDomains

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CountRecipientExternalDomains")]
        public void 宛先の外部ドメイン数取得_外部ドメイン数0()
        {
            var internalDomainList = new List<InternalDomain>();

            var testDisplayNameAndRecipient = new DisplayNameAndRecipient
            {
                All =
                {
                    ["miyake@noraneko.co.jp"] = "Takafumi Miyake",
                    ["test@noraneko.co.jp"] = "test",
                    ["test2@noraneko.co.jp"] = "test2"
                }
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testDisplayNameAndRecipient, "@noraneko.co.jp", internalDomainList, false };
            var result = (int)privateObject.Invoke("CountRecipientExternalDomains", args);

            Assert.AreEqual(result, 0);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CountRecipientExternalDomains")]
        public void 宛先の外部ドメイン数取得_外部ドメイン数2()
        {
            var internalDomainList = new List<InternalDomain>();

            var testDisplayNameAndRecipient = new DisplayNameAndRecipient
            {
                All =
                {
                    ["miyake@noraneko.co.jp"] = "Takafumi Miyake",
                    ["sample@sample.com"] = "サンプル太郎",
                    ["sample@sample2.com"] = "サンプル次郎"
                }
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testDisplayNameAndRecipient, "@noraneko.co.jp", internalDomainList, false };
            var result = (int)privateObject.Invoke("CountRecipientExternalDomains", args);

            Assert.AreEqual(result, 2);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CountRecipientExternalDomains")]
        public void 宛先の外部ドメイン数取得_外部ドメイン数5_内部ドメイン設定2()
        {
            var internalDomainList = new List<InternalDomain>
            {
                new InternalDomain {Domain = "@sample4.com"},
                new InternalDomain {Domain = "@sample5.com"}
            };

            var testDisplayNameAndRecipient = new DisplayNameAndRecipient
            {
                All =
                {
                    ["miyake@noraneko.co.jp"] = "Takafumi Miyake",
                    ["sample@sample.com"] = "サンプル太郎",
                    ["sample@sample2.com"] = "サンプル次郎",
                    ["sample@sample3.com"] = "サンプル三郎",
                    ["sample@sample4.com"] = "サンプル四郎",
                    ["sample@sample5.com"] = "サンプル五郎"
                }
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testDisplayNameAndRecipient, "@noraneko.co.jp", internalDomainList, false };
            var result = (int)privateObject.Invoke("CountRecipientExternalDomains", args);

            Assert.AreEqual(result, 3);
        }

        #endregion

        #region CalcDeferredMinutes

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CalcDeferredMinutes")]
        public void 送信保留する時間を算出_個別該当なし_全体設定あり()
        {
            var testDisplayNameAndRecipient = new DisplayNameAndRecipient
            {
                To =
                {
                    ["miyake@noraneko.co.jp"] = "Takafumi Miyake",
                    ["sample@sample.com"] = "サンプル太郎",
                    ["sample@sample2.com"] = "サンプル次郎"
                }
            };
            const int deferredMinutes = 5;
            var testDeferredDeliveryMinutes = new List<DeferredDeliveryMinutes>
            {
                new DeferredDeliveryMinutes {TargetAddress = "@", DeferredMinutes = deferredMinutes}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testDisplayNameAndRecipient, testDeferredDeliveryMinutes, false, 2 };
            var result = (int)privateObject.Invoke("CalcDeferredMinutes", args);

            Assert.AreEqual(result, deferredMinutes);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CalcDeferredMinutes")]
        public void 送信保留する時間を算出_個別該当あり_全体設定あり()
        {
            var testDisplayNameAndRecipient = new DisplayNameAndRecipient
            {
                To =
                {
                    ["miyake@noraneko.co.jp"] = "Takafumi Miyake",
                    ["sample@sample.com"] = "サンプル太郎 (sample@sample.com)",
                    ["sample@sample2.com"] = "サンプル次郎"
                }
            };
            const int targetDeferredMinutes = 10;
            var testDeferredDeliveryMinutes = new List<DeferredDeliveryMinutes>
            {
                new DeferredDeliveryMinutes {TargetAddress = "@", DeferredMinutes = 3},
                new DeferredDeliveryMinutes {TargetAddress = "@sample.com", DeferredMinutes = 5},
                new DeferredDeliveryMinutes {TargetAddress = "sample@sample2.com", DeferredMinutes = targetDeferredMinutes}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testDisplayNameAndRecipient, testDeferredDeliveryMinutes, false, 2 };
            var result = (int)privateObject.Invoke("CalcDeferredMinutes", args);

            Assert.AreEqual(result, targetDeferredMinutes);
        }

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("CalcDeferredMinutes")]
        public void 送信保留する時間を算出_個別該当なし_全体設定あり_外部ドメイン0_外部ドメイン0の際の無効設定あり()
        {
            var testDisplayNameAndRecipient = new DisplayNameAndRecipient
            {
                To =
                {
                    ["miyake@noraneko.co.jp"] = "Takafumi Miyake",
                    ["sample@noraneko.co.jp"] = "サンプル太郎",
                    ["sample2@noraneko.co.jp"] = "サンプル次郎"
                },
                Cc = { ["miyake@noraneko.co.jp"] = "Takafumi Miyake" },
                Bcc = { ["info@noraneko.co.jp"] = "Takafumi Miyake" },
            };

            var testDeferredDeliveryMinutes = new List<DeferredDeliveryMinutes>
            {
                new DeferredDeliveryMinutes {TargetAddress = "@", DeferredMinutes = 3}
            };

            var generateCheckList = new GenerateCheckList();
            var privateObject = new PrivateObject(generateCheckList);
            var args = new object[] { testDisplayNameAndRecipient, testDeferredDeliveryMinutes, true, 0 };
            var result = (int)privateObject.Invoke("CalcDeferredMinutes", args);

            Assert.AreEqual(result, 0);
        }

        #endregion

        #region GetExchangeDistributionListMembers

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetExchangeDistributionListMembers")]
        public void GetExchangeDistributionListMembers()
        {
            //TODO
        }

        #endregion

        #region GetContactGroupMembers

        [TestMethod, TestCategory("_GenerateCheckList"), TestCategory("GetContactGroupMembers")]
        public void GetContactGroupMembers()
        {
            //TODO
        }

        #endregion

        #endregion

        #region ThisAddIn

        //TODO

        #endregion
    }
}