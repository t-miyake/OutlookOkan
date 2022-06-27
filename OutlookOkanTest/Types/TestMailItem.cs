using Microsoft.Office.Interop.Outlook;
using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkanTest.Types
{
    public class TestMailItem : MailItem
    {
        public object Copy()
        {
            throw new NotImplementedException();
        }

        public void Delete()
        {
            throw new NotImplementedException();
        }

        public void Display(object Modal = null)
        {
            throw new NotImplementedException();
        }

        public object Move(Outlook.MAPIFolder DestFldr)
        {
            throw new NotImplementedException();
        }

        public void PrintOut()
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            throw new NotImplementedException();
        }

        public void SaveAs(string Path, object Type = null)
        {
            throw new NotImplementedException();
        }

        public void ClearConversationIndex()
        {
            throw new NotImplementedException();
        }

        public void ShowCategoriesDialog()
        {
            throw new NotImplementedException();
        }

        public void AddBusinessCard(Outlook.ContactItem contact)
        {
            throw new NotImplementedException();
        }

        public void MarkAsTask(Outlook.OlMarkInterval MarkInterval)
        {
            throw new NotImplementedException();
        }

        public void ClearTaskFlag()
        {
            throw new NotImplementedException();
        }

        public Outlook.Conversation GetConversation()
        {
            throw new NotImplementedException();
        }

        public Outlook.Application Application { get; }
        public Outlook.OlObjectClass Class { get; }
        public Outlook.NameSpace Session { get; }
        public object Parent { get; }
        public Outlook.Actions Actions { get; }
        public Outlook.Attachments Attachments { get; set; }
        public string BillingInformation { get; set; }
        public string Body { get; set; }
        public string Categories { get; set; }
        public string Companies { get; set; }
        public string ConversationIndex { get; }
        public string ConversationTopic { get; }
        public DateTime CreationTime { get; set; }
        public string EntryID { get; }
        public Outlook.FormDescription FormDescription { get; }
        public Outlook.Inspector GetInspector { get; }
        public Outlook.OlImportance Importance { get; set; }
        public DateTime LastModificationTime { get; }
        public object MAPIOBJECT { get; }
        public string MessageClass { get; set; }
        public string Mileage { get; set; }
        public bool NoAging { get; set; }
        public int OutlookInternalVersion { get; }
        public string OutlookVersion { get; }
        public bool Saved { get; }
        public Outlook.OlSensitivity Sensitivity { get; set; }
        public int Size { get; }
        public string Subject { get; set; }
        public bool UnRead { get; set; }
        public Outlook.UserProperties UserProperties { get; }
        public bool AlternateRecipientAllowed { get; set; }
        public bool AutoForwarded { get; set; }
        public string BCC { get; set; }
        public string CC { get; set; }
        public DateTime DeferredDeliveryTime { get; set; }
        public bool DeleteAfterSubmit { get; set; }
        public DateTime ExpiryTime { get; set; }
        public DateTime FlagDueBy { get; set; }
        public string FlagRequest { get; set; }
        public Outlook.OlFlagStatus FlagStatus { get; set; }
        public string HTMLBody { get; set; }
        public bool OriginatorDeliveryReportRequested { get; set; }
        public bool ReadReceiptRequested { get; set; }
        public string ReceivedByEntryID { get; }
        public string ReceivedByName { get; }
        public string ReceivedOnBehalfOfEntryID { get; }
        public string ReceivedOnBehalfOfName { get; }
        public DateTime ReceivedTime { get; }
        public bool RecipientReassignmentProhibited { get; set; }
        public Outlook.Recipients Recipients { get; }
        public bool ReminderOverrideDefault { get; set; }
        public bool ReminderPlaySound { get; set; }
        public bool ReminderSet { get; set; }
        public string ReminderSoundFile { get; set; }
        public DateTime ReminderTime { get; set; }
        public Outlook.OlRemoteStatus RemoteStatus { get; set; }
        public string ReplyRecipientNames { get; }
        public Outlook.Recipients ReplyRecipients { get; }
        public Outlook.MAPIFolder SaveSentMessageFolder { get; set; }
        public string SenderName { get; }
        public bool Sent { get; }
        public DateTime SentOn { get; }
        public string SentOnBehalfOfName { get; set; }
        public bool Submitted { get; }
        public string To { get; set; }
        public string VotingOptions { get; set; }
        public string VotingResponse { get; set; }
        public Outlook.Links Links { get; }
        public Outlook.ItemProperties ItemProperties { get; }
        public Outlook.OlBodyFormat BodyFormat { get; set; }
        public Outlook.OlDownloadState DownloadState { get; }
        public int InternetCodepage { get; set; }
        public Outlook.OlRemoteStatus MarkForDownload { get; set; }
        public bool IsConflict { get; }
        public bool IsIPFax { get; set; }
        public Outlook.OlFlagIcon FlagIcon { get; set; }
        public bool HasCoverSheet { get; set; }
        public bool AutoResolvedWinner { get; }
        public Outlook.Conflicts Conflicts { get; }
        public string SenderEmailAddress { get; set; }
        public string SenderEmailType { get; set; }
        public bool EnableSharedAttachments { get; set; }
        public Outlook.OlPermission Permission { get; set; }
        public Outlook.OlPermissionService PermissionService { get; set; }
        public Outlook.PropertyAccessor PropertyAccessor { get; }
        public Outlook.Account SendUsingAccount { get; set; }
        public string TaskSubject { get; set; }
        public DateTime TaskDueDate { get; set; }
        public DateTime TaskStartDate { get; set; }
        public DateTime TaskCompletedDate { get; set; }
        public DateTime ToDoTaskOrdinal { get; set; }
        public bool IsMarkedAsTask { get; }
        public string ConversationID { get; }
        public Outlook.AddressEntry Sender { get; set; }
        public string PermissionTemplateGuid { get; set; }
        public object RTFBody { get; set; }
        public string RetentionPolicyName { get; }
        public DateTime RetentionExpirationDate { get; }
        public event Outlook.ItemEvents_10_OpenEventHandler Open;
        public event Outlook.ItemEvents_10_CustomActionEventHandler CustomAction;
        public event Outlook.ItemEvents_10_CustomPropertyChangeEventHandler CustomPropertyChange;
        public event Outlook.ItemEvents_10_ForwardEventHandler Forward;
        public event Outlook.ItemEvents_10_CloseEventHandler Close;
        public event Outlook.ItemEvents_10_PropertyChangeEventHandler PropertyChange;
        public event Outlook.ItemEvents_10_ReadEventHandler Read;
        public event Outlook.ItemEvents_10_ReplyEventHandler Reply;
        public event Outlook.ItemEvents_10_ReplyAllEventHandler ReplyAll;
        public event Outlook.ItemEvents_10_SendEventHandler Send;
        public event Outlook.ItemEvents_10_WriteEventHandler Write;
        public event Outlook.ItemEvents_10_BeforeCheckNamesEventHandler BeforeCheckNames;
        public event Outlook.ItemEvents_10_AttachmentAddEventHandler AttachmentAdd;
        public event Outlook.ItemEvents_10_AttachmentReadEventHandler AttachmentRead;
        public event Outlook.ItemEvents_10_BeforeAttachmentSaveEventHandler BeforeAttachmentSave;
        public event Outlook.ItemEvents_10_BeforeDeleteEventHandler BeforeDelete;
        public event Outlook.ItemEvents_10_AttachmentRemoveEventHandler AttachmentRemove;
        public event Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler BeforeAttachmentAdd;
        public event Outlook.ItemEvents_10_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreview;
        public event Outlook.ItemEvents_10_BeforeAttachmentReadEventHandler BeforeAttachmentRead;
        public event Outlook.ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFile;
        public event Outlook.ItemEvents_10_UnloadEventHandler Unload;
        public event Outlook.ItemEvents_10_BeforeAutoSaveEventHandler BeforeAutoSave;
        public event Outlook.ItemEvents_10_BeforeReadEventHandler BeforeRead;
        public event Outlook.ItemEvents_10_AfterWriteEventHandler AfterWrite;
        public event Outlook.ItemEvents_10_ReadCompleteEventHandler ReadComplete;

        void _MailItem.Close(OlInspectorClose SaveMode)
        {
            throw new NotImplementedException();
        }

        MailItem _MailItem.Forward()
        {
            throw new NotImplementedException();
        }

        MailItem _MailItem.Reply()
        {
            throw new NotImplementedException();
        }

        MailItem _MailItem.ReplyAll()
        {
            throw new NotImplementedException();
        }

        void _MailItem.Send()
        {
            throw new NotImplementedException();
        }
    }
}