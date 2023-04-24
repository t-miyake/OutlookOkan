using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OutlookOkan.Handlers
{
    internal static class MailHeaderHandler
    {
        internal static Dictionary<string, string> ValidateEmailHeader(string emailHeader)
        {
            var results = new Dictionary<string, string>
            {
                ["SPF"] = "NONE",
                ["SPF IP"] = "NONE",
                ["DKIM"] = "NONE",
                ["DKIM Domain"] = "NONE",
                ["DMARC"] = "NONE"
            };

            // SPF Validation
            var spfRegex = new Regex(@"Received-SPF:\s*(?<result>pass|fail|softfail|neutral|temperror|permerror|none).*\b(does\s+not\s+)?designate[s]?\s+(?<ip>[^ ]+)\s+as", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var spfMatch = spfRegex.Match(emailHeader);
            if (spfMatch.Success)
            {
                results["SPF"] = spfMatch.Groups["result"].Value.ToUpper();
                results["SPF IP"] = spfMatch.Groups["ip"].Value;
            }
            else
            {
                results["SPF"] = "NULL";
            }

            // DKIM Validation
            var dkimRegex = new Regex(@"Authentication-Results:.*?dkim=(?<result>pass|policy|fail|softfail|hardfail|neutral|temperror|permerror|none).*?header.(d|i)=(?<domain>[^(;| )]+)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dkimMatch = dkimRegex.Match(emailHeader);
            if (dkimMatch.Success)
            {
                results["DKIM"] = dkimMatch.Groups["result"].Value.ToUpper();
                results["DKIM Domain"] = dkimMatch.Groups["domain"].Value;
            }

            // DMARC Validation
            var dmarcRegex = new Regex(@"Authentication-Results:.*?dmarc=(?<result>pass|bestguesspass|softfail|fail|none)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dmarcMatch = dmarcRegex.Match(emailHeader);
            if (dmarcMatch.Success)
            {
                results["DMARC"] = dmarcMatch.Groups["result"].Value.ToUpper();
            }

            return results;
        }
    }
}