using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookOkan.Handlers
{
    /// <summary>
    /// メールヘッダの解析を行う
    /// </summary>
    internal static class MailHeaderHandler
    {
        /// <summary>
        /// メールヘッダを解析し、SPF、DKIM、DMARCなどの検証結果を返す
        /// </summary>
        /// <param name="emailHeader">メールヘッダ</param>
        /// <returns>解析結果</returns>
        internal static Dictionary<string, string> ValidateEmailHeader(string emailHeader)
        {
            var results = new Dictionary<string, string>
            {
                ["From Domain"] = "NONE",
                ["ReturnPath Domain"] = "NONE",
                ["SPF"] = "NONE",
                ["SPF IP"] = "NONE",
                ["SPF Alignment"] = "NONE",
                ["DKIM"] = "NONE",
                ["DKIM Domain"] = "NONE",
                ["DKIM Alignment"] = "NONE",
                ["DMARC"] = "NONE",
                ["Internal"] = "FALSE"
            };

            if (string.IsNullOrEmpty(emailHeader))
            {
                return null;
            }

            if (IsInternalMail(emailHeader))
            {
                results["Internal"] = "TRUE";
            }

            var fromDomain = string.Empty;
            var fromRegex = new Regex(@"^From:\s*.*(?:\r?\n\s+.*)*", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            var fromMatch = fromRegex.Match(emailHeader);
            if (fromMatch.Success)
            {
                var fromHeader = fromMatch.Value;
                var domainRegex = new Regex(@"<.*?@(?<domain>[^\s>]+)>", RegexOptions.IgnoreCase);
                var domainMatch = domainRegex.Match(fromHeader);

                if (!domainMatch.Success)
                {
                    var alternativeDomainRegex = new Regex(@"[^<\s]+@(?<domain>[^\s>]+)", RegexOptions.IgnoreCase);
                    domainMatch = alternativeDomainRegex.Match(fromHeader);
                }

                fromDomain = domainMatch.Success ? domainMatch.Groups["domain"].Value : string.Empty;
                results["From Domain"] = fromDomain;
            }

            // SPF検証
            var spfRegex = new Regex(@"Received-SPF:\s*(?<result>pass|fail|softfail|neutral|temperror|permerror|none).*\b(does\s+not\s+)?designate[s]?\s+(?<ip>[^ ]+)\s+as", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var spfMatch = spfRegex.Match(emailHeader);
            if (spfMatch.Success)
            {
                results["SPF"] = spfMatch.Groups["result"].Value.ToUpper();
                results["SPF IP"] = spfMatch.Groups["ip"].Value;
            }

            // SPFアライメント検証
            var returnPathRegex = new Regex(@"Return-Path:\s*.*@(?<domain>[^\s>]+)");
            var returnPathMatch = returnPathRegex.Match(emailHeader);
            if (returnPathMatch.Success && fromDomain != string.Empty)
            {
                var returnPathDomain = returnPathMatch.Groups["domain"].Value;
                results["ReturnPath Domain"] = returnPathDomain;
                results["SPF Alignment"] = returnPathDomain.Equals(fromDomain, StringComparison.OrdinalIgnoreCase) || returnPathDomain.ToLower().Contains(fromDomain.ToLower()) || fromDomain.ToLower().Contains(returnPathDomain.ToLower()) ? "PASS" : "FAIL";
            }

            // DKIM検証
            var dkimRegex = new Regex(@"Authentication-Results:.*?dkim=(?<result>pass|policy|fail|softfail|hardfail|neutral|temperror|permerror|none).*?header.d=(?<domain>[^(;| )]+)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dkimMatch = dkimRegex.Match(emailHeader);
            if (dkimMatch.Success)
            {
                results["DKIM"] = dkimMatch.Groups["result"].Value.ToUpper();
            }

            // DKIMアライメント検証
            var dkimSignatureRegex = new Regex(@"DKIM-Signature:.*?d=(?<domain>[^(;| )]+)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dkimMatches = dkimSignatureRegex.Matches(emailHeader);
            var dkimAlignmentPass = false;
            var dkimDomains = new List<string>();

            foreach (Match match in dkimMatches)
            {
                var dkimDomain = match.Groups["domain"].Value;
                if (string.IsNullOrEmpty(dkimDomain)) continue;

                dkimDomains.Add(dkimDomain);
                if (dkimDomain.Equals(fromDomain, StringComparison.OrdinalIgnoreCase) ||
                    dkimDomain.ToLower().Contains(fromDomain.ToLower()) ||
                    fromDomain.ToLower().Contains(dkimDomain.ToLower()))
                {
                    dkimAlignmentPass = true;
                }
            }
            results["DKIM Domain"] = string.Join(", ", dkimDomains);
            results["DKIM Alignment"] = dkimAlignmentPass ? "PASS" : "FAIL";

            // DMARC検証
            var dmarcRegex = new Regex(@"Authentication-Results:.*?dmarc=(?<result>pass|bestguesspass|softfail|fail|none)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dmarcMatch = dmarcRegex.Match(emailHeader);
            if (dmarcMatch.Success)
            {
                results["DMARC"] = dmarcMatch.Groups["result"].Value.ToUpper();
            }

            return results;
        }

        /// <summary>
        /// DMARCの検証結果を独自判定する
        /// </summary>
        /// <param name="spfResult"></param>
        /// <param name="spfAlignmentResult"></param>
        /// <param name="dkimResult"></param>
        /// <param name="dkimAlignmentResult"></param>
        /// <returns>DMARCの検証結果</returns>
        public static string DetermineDmarcResult(string spfResult, string spfAlignmentResult, string dkimResult, string dkimAlignmentResult)
        {
            if (string.IsNullOrEmpty(spfResult) || string.IsNullOrEmpty(spfAlignmentResult) || string.IsNullOrEmpty(dkimResult) || string.IsNullOrEmpty(dkimAlignmentResult))
            {
                return "FAIL";
            }

            // NONEはFAILとして扱う
            spfResult = spfResult.ToUpper() == "NONE" ? "FAIL" : spfResult.ToUpper();
            spfAlignmentResult = spfAlignmentResult.ToUpper() == "NONE" ? "FAIL" : spfAlignmentResult.ToUpper();
            dkimResult = dkimResult.ToUpper() == "NONE" ? "FAIL" : dkimResult.ToUpper();
            dkimAlignmentResult = dkimAlignmentResult.ToUpper() == "NONE" ? "FAIL" : dkimAlignmentResult.ToUpper();

            var key = $"{spfResult}_{spfAlignmentResult}_{dkimResult}_{dkimAlignmentResult}";

            //SPF認証_SPFアライメント_DKIM認証_DKIMアライメント
            var dmarcResults = new Dictionary<string, string>
            {
                { "PASS_PASS_PASS_PASS", "PASS" }, // 両方の認証とアライメントが成功
                { "PASS_PASS_PASS_FAIL", "PASS" }, // SPFの認証とアライメントが成功、DKIMの認証が成功
                { "PASS_PASS_FAIL_PASS", "PASS" }, // SPFの認証とアライメントが成功、DKIMのアライメントが成功
                { "PASS_PASS_FAIL_FAIL", "PASS" }, // SPFの認証とアライメントが成功
                { "PASS_FAIL_PASS_PASS", "PASS" }, // SPFの認証が成功、DKIMの認証とアライメントが成功
                { "FAIL_PASS_PASS_PASS", "PASS" }, // SPFのアライメントが成功、DKIMの認証とアライメントが成功
                { "FAIL_FAIL_PASS_PASS", "PASS" }, // DKIMの認証とアライメントが成功
                
                { "PASS_FAIL_PASS_FAIL", "FAIL" }, // SPFの認証が成功、DKIMの認証が成功
                { "PASS_FAIL_FAIL_PASS", "FAIL" }, // SPFの認証が成功、DKIMのアライメントが成功
                { "PASS_FAIL_FAIL_FAIL", "FAIL" }, // SPFの認証が成功
                { "FAIL_PASS_PASS_FAIL", "FAIL" }, // SPFのアライメントが成功、DKIMの認証が成功
                { "FAIL_PASS_FAIL_PASS", "FAIL" }, // SPFのアライメントが成功、DKIMのアライメントが成功
                { "FAIL_PASS_FAIL_FAIL", "FAIL" }, // SPFのアライメントが成功
                { "FAIL_FAIL_PASS_FAIL", "FAIL" }, // DKIMの認証が成功
                { "FAIL_FAIL_FAIL_PASS", "FAIL" }, // DKIMのアライメントが成功
                { "FAIL_FAIL_FAIL_FAIL", "FAIL" }  // すべて失敗
            };
            return dmarcResults.TryGetValue(key, out var result) ? result : "FAIL";
        }

        /// <summary>
        /// 内部メールか否かの判定
        /// </summary>
        /// <param name="emailHeader">メールヘッダ</param>
        /// <returns>判定結果</returns>
        internal static bool IsInternalMail(string emailHeader)
        {
            // Receivedヘッダをすべて取得
            var receivedRegex = new Regex(@"^Received:.*", RegexOptions.Multiline);
            var matches = receivedRegex.Matches(emailHeader);

            var receivedHeaders = (from Match match in matches select match.Value).ToList();

            // 受信ヘッダの数が多い場合は外部メールと判定
            if (receivedHeaders.Count > 3)
            {
                return false;
            }

            // 受信ヘッダが複数ある場合、連続したドメイン名が一致するかどうかを確認
            var domainRegex = new Regex(@"from\s([^\s]+)", RegexOptions.IgnoreCase);
            string previousDomain = null;

            foreach (var currentDomain in from header in receivedHeaders select domainRegex.Match(header) into domainMatch where domainMatch.Success select ExtractMainDomain(domainMatch.Groups[1].Value))
            {
                if (previousDomain != null && previousDomain.Equals(currentDomain, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                previousDomain = currentDomain;
            }

            return false;
        }

        private static string ExtractMainDomain(string domain)
        {
            var parts = domain.Split('.');
            var length = parts.Length;

            return length > 2 ? string.Join(".", parts.Skip(length - 3)) : domain;
        }
    }
}