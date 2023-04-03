using OutlookOkan.Types;
using System.Collections.ObjectModel;
using System.Linq;

namespace OutlookOkan.ViewModels
{
    internal sealed class ConfirmationWindowViewModel : ViewModelBase
    {
        internal ConfirmationWindowViewModel(CheckList checkList)
        {
            _checkList = checkList;

            GenerateAlertsCollection();
            GenerateAddressesCollection();
            GenerateAttachmentsCollection();

            UpDateItemsCount();

            //画面サイズを自動で変更させるための処置
            if (ToAddresses.Count >= 5)
            {
                var tempCount = ToAddresses.Count;
                tempCount -= 4;
                for (var i = 1; i < tempCount; i++)
                {
                    if (i == 7) break;

                    ToGridHeight += AddTextHeight;
                    MailBodyTextBoxHeight += AddTextHeight;
                    WindowHeight += AddTextHeight;
                }
            }
            if (CcAddresses.Count >= 5)
            {
                var tempCount = CcAddresses.Count;
                tempCount -= 4;
                for (var i = 1; i < tempCount; i++)
                {
                    if (i == 7) break;

                    CcGridHeight += AddTextHeight;
                    MailBodyTextBoxHeight += AddTextHeight;
                    WindowHeight += AddTextHeight;
                }
            }
            if (BccAddresses.Count >= 5)
            {
                var tempCount = BccAddresses.Count;
                tempCount -= 4;
                for (var i = 1; i < tempCount; i++)
                {
                    if (i == 7) break;

                    BccGridHeight += AddTextHeight;
                    MailBodyTextBoxHeight += AddTextHeight;
                    WindowHeight += AddTextHeight;
                }
            }

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
            ToAddressCount = ToAddresses.Count.ToString();
            CcAddressCount = CcAddresses.Count.ToString();
            BccAddressCount = BccAddresses.Count.ToString();

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
                OnPropertyChanged(nameof(Alerts));
            }
        }

        private ObservableCollection<Address> _toAddresses = new ObservableCollection<Address>();
        public ObservableCollection<Address> ToAddresses
        {
            get => _toAddresses;
            set
            {
                _toAddresses = value;
                OnPropertyChanged(nameof(ToAddresses));
            }
        }

        private ObservableCollection<Address> _ccAddresses = new ObservableCollection<Address>();
        public ObservableCollection<Address> CcAddresses
        {
            get => _ccAddresses;
            set
            {
                _ccAddresses = value;
                OnPropertyChanged(nameof(CcAddresses));
            }
        }

        private ObservableCollection<Address> _bccAddresses = new ObservableCollection<Address>();
        public ObservableCollection<Address> BccAddresses
        {
            get => _bccAddresses;
            set
            {
                _bccAddresses = value;
                OnPropertyChanged(nameof(BccAddresses));
            }
        }

        private ObservableCollection<Attachment> _attachments = new ObservableCollection<Attachment>();
        public ObservableCollection<Attachment> Attachments
        {
            get => _attachments;
            set
            {
                _attachments = value;
                OnPropertyChanged(nameof(Attachments));
            }
        }

        private bool _isCanSendMail;
        public bool IsCanSendMail
        {
            get => _isCanSendMail;
            set
            {
                _isCanSendMail = value;
                OnPropertyChanged(nameof(IsCanSendMail));
            }
        }

        private string _addressCount = Properties.Resources.DestinationEmailaddress + " ()";

        public string AddressCount
        {
            get => _addressCount;
            set
            {
                _addressCount = value;
                OnPropertyChanged(nameof(AddressCount));
            }
        }

        private string _toAddressCount = "To ()";
        public string ToAddressCount
        {
            get => _toAddressCount;
            set
            {
                _toAddressCount = value;
                OnPropertyChanged(nameof(ToAddressCount));
            }
        }

        private string _ccAddressCount = "Cc ()";
        public string CcAddressCount
        {
            get => _ccAddressCount;
            set
            {
                _ccAddressCount = value;
                OnPropertyChanged(nameof(CcAddressCount));
            }
        }

        private string _bccAddressCount = "Bcc ()";
        public string BccAddressCount
        {
            get => _bccAddressCount;
            set
            {
                _bccAddressCount = value;
                OnPropertyChanged(nameof(BccAddressCount));
            }
        }

        private string _alertCount = Properties.Resources.ImportantAlert + " ()";
        public string AlertCount
        {
            get => _alertCount;
            set
            {
                _alertCount = value;
                OnPropertyChanged(nameof(AlertCount));
            }
        }

        private string _attachmentCount = Properties.Resources.Attachments + " ()";
        public string AttachmentCount
        {
            get => _attachmentCount;
            set
            {
                _attachmentCount = value;
                OnPropertyChanged(nameof(AttachmentCount));
            }
        }

        private int _deferredDeliveryMinutes;
        public int DeferredDeliveryMinutes
        {
            get => _deferredDeliveryMinutes;
            set
            {
                _deferredDeliveryMinutes = value;
                OnPropertyChanged(nameof(DeferredDeliveryMinutes));
            }
        }

        public string Sender => _checkList.Sender;
        public string Subject => _checkList.Subject;
        public string MailType => _checkList.MailType;
        public string MailBody => _checkList.MailBody;


        private int _windowHeight = 590;
        public int WindowHeight
        {
            get => _windowHeight;
            set
            {
                _windowHeight = value;
                OnPropertyChanged(nameof(WindowHeight));
            }
        }

        private int _toGridHeight = 82;
        public int ToGridHeight
        {
            get => _toGridHeight;
            set
            {
                _toGridHeight = value;
                OnPropertyChanged(nameof(ToGridHeight));
            }
        }

        private int _ccGridHeight = 82;
        public int CcGridHeight
        {
            get => _ccGridHeight;
            set
            {
                _ccGridHeight = value;
                OnPropertyChanged(nameof(CcGridHeight));
            }
        }

        private int _bccGridHeight = 82;
        public int BccGridHeight
        {
            get => _bccGridHeight;
            set
            {
                _bccGridHeight = value;
                OnPropertyChanged(nameof(BccGridHeight));
            }
        }

        private int _mailBodyTextBoxHeight = 79;
        public int MailBodyTextBoxHeight
        {
            get => _mailBodyTextBoxHeight;
            set
            {
                _mailBodyTextBoxHeight = value;
                OnPropertyChanged(nameof(MailBodyTextBoxHeight));
            }
        }

        private const int AddTextHeight = 18;
    }
}