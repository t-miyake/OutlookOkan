using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkanTest.Types
{
    public class TestRecipient : Outlook.Recipient
    {
        public void Delete()
        {
            throw new NotImplementedException();
        }

        public string FreeBusy(DateTime Start, int MinPerChar, object CompleteFormat = null)
        {
            throw new NotImplementedException();
        }

        public bool Resolve()
        {
            throw new NotImplementedException();
        }

        public Outlook.Application Application { get; }
        public Outlook.OlObjectClass Class { get; }
        public Outlook.NameSpace Session { get; }
        public object Parent { get; }
        public string Address { get; }
        public Outlook.AddressEntry AddressEntry { get; set; }
        public string AutoResponse { get; set; }
        public Outlook.OlDisplayType DisplayType { get; }
        public string EntryID { get; }
        public int Index { get; }
        public Outlook.OlResponseStatus MeetingResponseStatus { get; }
        public string Name { get; set; }
        public bool Resolved { get; }
        public Outlook.OlTrackingStatus TrackingStatus { get; set; }
        public DateTime TrackingStatusTime { get; set; }
        public int Type { get; set; }
        public Outlook.PropertyAccessor PropertyAccessor { get; }
        public bool Sendable { get; set; }
    }
}