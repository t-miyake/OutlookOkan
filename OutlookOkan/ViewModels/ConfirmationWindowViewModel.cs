using OutlookOkan.Types;
using System.Collections.ObjectModel;
using System.Linq;

namespace OutlookOkan.ViewModels
{
    public class ConfirmationWindowViewModel : ViewModelBase
    {
        public ConfirmationWindowViewModel(CheckList checkList)
        {
            _checkList = checkList;

            GenerateAlertsCollection();
            GenerateAddressesCollection();
            GenerateAttachmentsCollection();

            UpDateItemsCount();

            ToggleSendButton();
        }

        private void GenerateAlertsCollection()
        {
            foreach (var alert in _checkList.Alerts)
            {
                Alerts.Add(alert);
            }

            if (Alerts.Count == 0)
            {
                Alerts.Add(new Alert { AlertMessage = Properties.Resources.NoAlert, IsChecked = true, IsWhite = true, IsImportant = false });
            }
        }

        private void GenerateAddressesCollection()
        {
            foreach (var address in _checkList.ToAddresses)
            {
                ToAddresses.Add(address);
            }

            foreach (var address in _checkList.CcAddresses)
            {
                CcAddresses.Add(address);
            }

            foreach (var address in _checkList.BccAddresses)
            {
                BccAddresses.Add(address);
            }
        }

        private void GenerateAttachmentsCollection()
        {
            foreach (var attachment in _checkList.Attachments)
            {
                Attachments.Add(attachment);
            }
        }

        private void UpDateItemsCount()
        {
            AlertCount = Properties.Resources.ImportantAlert + " (" + Alerts.Count + ")";

            AddressCount = Properties.Resources.DestinationEmailaddress + " (" + (ToAddresses.Count + CcAddresses.Count + BccAddresses.Count) + ")";
            ToAddressCount = "To (" + ToAddresses.Count + ")";
            CcAddressCount = "CC (" + CcAddresses.Count + ")";
            BccAddressCount = "BCC (" + BccAddresses.Count + ")";

            AttachmentCount = Properties.Resources.Attachments + " (" + Attachments.Count + ")";
        }

        public void ToggleSendButton()
        {
            var isToAddressesCompleteChecked = ToAddresses.Count(x => x.IsChecked) == ToAddresses.Count;
            var isCcAddressesCompleteChecked = CcAddresses.Count(x => x.IsChecked) == CcAddresses.Count;
            var isBccAddressesCompleteChecked = BccAddresses.Count(x => x.IsChecked) == BccAddresses.Count;
            var isAlertsCompleteChecked = Alerts.Count(x => x.IsChecked) == Alerts.Count;
            var isAttachmentsCompleteChecked = Attachments.Count(x => x.IsChecked) == Attachments.Count;

            if (isToAddressesCompleteChecked && isCcAddressesCompleteChecked && isBccAddressesCompleteChecked &&
                isAlertsCompleteChecked && isAttachmentsCompleteChecked)
            {
                IsCanSendMail = true;
            }
            else
            {
                IsCanSendMail = false;
            }
        }

        private readonly CheckList _checkList;

        private ObservableCollection<Alert> _alerts = new ObservableCollection<Alert>();
        public ObservableCollection<Alert> Alerts
        {
            get => _alerts;
            set
            {
                _alerts = value;
                OnPropertyChanged("Alerts");
            }
        }

        private ObservableCollection<Address> _toAddresses = new ObservableCollection<Address>();
        public ObservableCollection<Address> ToAddresses
        {
            get => _toAddresses;
            set
            {
                _toAddresses = value;
                OnPropertyChanged("ToAddresses");
            }
        }

        private ObservableCollection<Address> _ccAddresses = new ObservableCollection<Address>();
        public ObservableCollection<Address> CcAddresses
        {
            get => _ccAddresses;
            set
            {
                _ccAddresses = value;
                OnPropertyChanged("CcAddresses");
            }
        }

        private ObservableCollection<Address> _bccAddresses = new ObservableCollection<Address>();
        public ObservableCollection<Address> BccAddresses
        {
            get => _bccAddresses;
            set
            {
                _bccAddresses = value;
                OnPropertyChanged("BccAddresses");
            }
        }

        private ObservableCollection<Attachment> _attachments = new ObservableCollection<Attachment>();
        public ObservableCollection<Attachment> Attachments
        {
            get => _attachments;
            set
            {
                _attachments = value;
                OnPropertyChanged("Attachments");
            }
        }

        private bool _isCanSendMail;
        public bool IsCanSendMail
        {
            get => _isCanSendMail;
            set
            {
                _isCanSendMail = value;
                OnPropertyChanged("IsCanSendMail");
            }
        }

        private string _addressCount = Properties.Resources.DestinationEmailaddress + " ()";

        public string AddressCount
        {
            get => _addressCount;
            set
            {
                _addressCount = value;
                OnPropertyChanged("AddressCount");
            }
        }

        private string _toAddressCount = "To ()";
        public string ToAddressCount
        {
            get => _toAddressCount;
            set
            {
                _toAddressCount = value;
                OnPropertyChanged("ToAddressCount");
            }
        }

        private string _ccAddressCount = "CC ()";
        public string CcAddressCount
        {
            get => _ccAddressCount;
            set
            {
                _ccAddressCount = value;
                OnPropertyChanged("CcAddressCount");
            }
        }

        private string _bccAddressCount = "BCC ()";
        public string BccAddressCount
        {
            get => _bccAddressCount;
            set
            {
                _bccAddressCount = value;
                OnPropertyChanged("BccAddressCount");
            }
        }

        private string _alertCount = Properties.Resources.ImportantAlert + " ()";
        public string AlertCount
        {
            get => _alertCount;
            set
            {
                _alertCount = value;
                OnPropertyChanged("AlertCount");
            }
        }

        private string _attachmentCount = Properties.Resources.Attachments + " ()";
        public string AttachmentCount
        {
            get => _attachmentCount;
            set
            {
                _attachmentCount = value;
                OnPropertyChanged("AttachmentCount");
            }
        }

        public string Sender => _checkList.Sender;
        public string Subject => _checkList.Subject;
        public string MailType => _checkList.MailType;
        public string MailBody => _checkList.MailBody;
    }
}