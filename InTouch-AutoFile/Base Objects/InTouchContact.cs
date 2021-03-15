using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace InTouch_AutoFile
{
    /// <summary>
    /// An extention of the Outlook ContactItem.
    /// </summary>
    public sealed class InTouchContact : ContactItem
    {
        #region Base Object

        private readonly ContactItem contact;

        public void Close(OlInspectorClose SaveMode)
        {
            SaveData();
            contact.Close(SaveMode);
        }

        public dynamic Copy()
        {
            return contact.Copy();
        }

        public void Delete()
        {
            contact.Delete();
        }

        public void Display(object Modal)
        {
            contact.Display(Modal);
        }

        public dynamic Move(MAPIFolder DestFldr)
        {
            return contact.Move(DestFldr);
        }

        public void PrintOut()
        {
            contact.PrintOut();
        }

        public void Save()
        {            
            contact.Save();
        }

        public void SaveAs(string Path, object Type)
        {
            SaveData();
            contact.SaveAs(Path, Type);
        }

        public MailItem ForwardAsVcard()
        {
            return contact.ForwardAsVcard();
        }

        public void ShowCategoriesDialog()
        {
            contact.ShowCategoriesDialog();
        }

        public void AddPicture(string Path)
        {
            contact.AddPicture(Path);
        }

        public void RemovePicture()
        {
            contact.RemovePicture();
        }

        public MailItem ForwardAsBusinessCard()
        {
            return contact.ForwardAsBusinessCard();
        }

        public void ShowBusinessCardEditor()
        {
            contact.ShowBusinessCardEditor();
        }

        public void SaveBusinessCardImage(string Path)
        {
            contact.SaveBusinessCardImage(Path);
        }

        public void ShowCheckPhoneDialog(OlContactPhoneNumber PhoneNumber)
        {
            contact.ShowCheckPhoneDialog(PhoneNumber);
        }

        public void MarkAsTask(OlMarkInterval MarkInterval)
        {
            contact.MarkAsTask(MarkInterval);
        }

        public void ClearTaskFlag()
        {
            contact.ClearTaskFlag();
        }

        public void ResetBusinessCard()
        {
            contact.ResetBusinessCard();
        }

        public void AddBusinessCardLogoPicture(string Path)
        {
            contact.AddBusinessCardLogoPicture(Path);
        }

        public Conversation GetConversation()
        {
            return contact.GetConversation();
        }

        public void ShowCheckFullNameDialog()
        {
            contact.ShowCheckFullNameDialog();
        }

        public void ShowCheckAddressDialog(OlMailingAddress MailingAddress)
        {
            contact.ShowCheckAddressDialog(MailingAddress);
        }

        public Application Application => contact.Application;

        public OlObjectClass Class => contact.Class;

        public NameSpace Session => contact.Session;

        public dynamic Parent => contact.Parent;

        public Actions Actions => contact.Actions;

        public Attachments Attachments => contact.Attachments;

        public string BillingInformation { get => contact.BillingInformation; set => contact.BillingInformation = value; }

        public string Body { get => contact.Body; set => contact.Body = value; }

        public string Categories { get => contact.Categories; set => contact.Categories = value; }

        public string Companies { get => contact.Companies; set => contact.Companies = value; }

        public string ConversationIndex => contact.ConversationIndex;

        public string ConversationTopic => contact.ConversationTopic;

        public DateTime CreationTime => contact.CreationTime;

        public string EntryID => contact.EntryID;

        public FormDescription FormDescription => contact.FormDescription;

        public Inspector GetInspector => contact.GetInspector;

        public OlImportance Importance { get => contact.Importance; set => contact.Importance = value; }

        public DateTime LastModificationTime => contact.LastModificationTime;

        public dynamic MAPIOBJECT => contact.MAPIOBJECT;

        public string MessageClass { get => contact.MessageClass; set => contact.MessageClass = value; }

        public string Mileage { get => contact.Mileage; set => contact.Mileage = value; }

        public bool NoAging { get => contact.NoAging; set => contact.NoAging = value; }

        public int OutlookInternalVersion => contact.OutlookInternalVersion;

        public string OutlookVersion => contact.OutlookVersion;

        public bool Saved => contact.Saved;

        public OlSensitivity Sensitivity { get => contact.Sensitivity; set => contact.Sensitivity = value; }

        public int Size => contact.Size;

        public string Subject { get => contact.Subject; set => contact.Subject = value; }

        public bool UnRead { get => contact.UnRead; set => contact.UnRead = value; }

        public UserProperties UserProperties => contact.UserProperties;

        public string Account { get => contact.Account; set => contact.Account = value; }

        public DateTime Anniversary { get => contact.Anniversary; set => contact.Anniversary = value; }

        public string AssistantName { get => contact.AssistantName; set => contact.AssistantName = value; }

        public string AssistantTelephoneNumber { get => contact.AssistantTelephoneNumber; set => contact.AssistantTelephoneNumber = value; }

        public DateTime Birthday { get => contact.Birthday; set => contact.Birthday = value; }

        public string Business2TelephoneNumber { get => contact.Business2TelephoneNumber; set => contact.Business2TelephoneNumber = value; }

        public string BusinessAddress { get => contact.BusinessAddress; set => contact.BusinessAddress = value; }

        public string BusinessAddressCity { get => contact.BusinessAddressCity; set => contact.BusinessAddressCity = value; }

        public string BusinessAddressCountry { get => contact.BusinessAddressCountry; set => contact.BusinessAddressCountry = value; }

        public string BusinessAddressPostalCode { get => contact.BusinessAddressPostalCode; set => contact.BusinessAddressPostalCode = value; }

        public string BusinessAddressPostOfficeBox { get => contact.BusinessAddressPostOfficeBox; set => contact.BusinessAddressPostOfficeBox = value; }

        public string BusinessAddressState { get => contact.BusinessAddressState; set => contact.BusinessAddressState = value; }

        public string BusinessAddressStreet { get => contact.BusinessAddressStreet; set => contact.BusinessAddressStreet = value; }

        public string BusinessFaxNumber { get => contact.BusinessFaxNumber; set => contact.BusinessFaxNumber = value; }

        public string BusinessHomePage { get => contact.BusinessHomePage; set => contact.BusinessHomePage = value; }

        public string BusinessTelephoneNumber { get => contact.BusinessTelephoneNumber; set => contact.BusinessTelephoneNumber = value; }

        public string CallbackTelephoneNumber { get => contact.CallbackTelephoneNumber; set => contact.CallbackTelephoneNumber = value; }

        public string CarTelephoneNumber { get => contact.CarTelephoneNumber; set => contact.CarTelephoneNumber = value; }

        public string Children { get => contact.Children; set => contact.Children = value; }

        public string CompanyAndFullName => contact.CompanyAndFullName;

        public string CompanyLastFirstNoSpace => contact.CompanyLastFirstNoSpace;

        public string CompanyLastFirstSpaceOnly => contact.CompanyLastFirstSpaceOnly;

        public string CompanyMainTelephoneNumber { get => contact.CompanyMainTelephoneNumber; set => contact.CompanyMainTelephoneNumber = value; }

        public string CompanyName { get => contact.CompanyName; set => contact.CompanyName = value; }

        public string ComputerNetworkName { get => contact.ComputerNetworkName; set => contact.ComputerNetworkName = value; }

        public string CustomerID { get => contact.CustomerID; set => contact.CustomerID = value; }

        public string Department { get => contact.Department; set => contact.Department = value; }

        public string Email1Address { get => contact.Email1Address; set => contact.Email1Address = value; }

        public string Email1AddressType { get => contact.Email1AddressType; set => contact.Email1AddressType = value; }

        public string Email1DisplayName { get => contact.Email1DisplayName; set => contact.Email1DisplayName = value; }

        public string Email1EntryID => contact.Email1EntryID;

        public string Email2Address { get => contact.Email2Address; set => contact.Email2Address = value; }

        public string Email2AddressType { get => contact.Email2AddressType; set => contact.Email2AddressType = value; }

        public string Email2DisplayName { get => contact.Email2DisplayName; set => contact.Email2DisplayName = value; }

        public string Email2EntryID => contact.Email2EntryID;

        public string Email3Address { get => contact.Email3Address; set => contact.Email3Address = value; }

        public string Email3AddressType { get => contact.Email3AddressType; set => contact.Email3AddressType = value; }

        public string Email3DisplayName { get => contact.Email3DisplayName; set => contact.Email3DisplayName = value; }

        public string Email3EntryID => contact.Email3EntryID;

        public string FileAs { get => contact.FileAs; set => contact.FileAs = value; }

        public string FirstName { get => contact.FirstName; set => contact.FirstName = value; }

        public string FTPSite { get => contact.FTPSite; set => contact.FTPSite = value; }

        public string FullName { get => contact.FullName; set => contact.FullName = value; }

        public string FullNameAndCompany => contact.FullNameAndCompany;

        public OlGender Gender { get => contact.Gender; set => contact.Gender = value; }

        public string GovernmentIDNumber { get => contact.GovernmentIDNumber; set => contact.GovernmentIDNumber = value; }

        public string Hobby { get => contact.Hobby; set => contact.Hobby = value; }

        public string Home2TelephoneNumber { get => contact.Home2TelephoneNumber; set => contact.Home2TelephoneNumber = value; }

        public string HomeAddress { get => contact.HomeAddress; set => contact.HomeAddress = value; }

        public string HomeAddressCity { get => contact.HomeAddressCity; set => contact.HomeAddressCity = value; }

        public string HomeAddressCountry { get => contact.HomeAddressCountry; set => contact.HomeAddressCountry = value; }

        public string HomeAddressPostalCode { get => contact.HomeAddressPostalCode; set => contact.HomeAddressPostalCode = value; }

        public string HomeAddressPostOfficeBox { get => contact.HomeAddressPostOfficeBox; set => contact.HomeAddressPostOfficeBox = value; }

        public string HomeAddressState { get => contact.HomeAddressState; set => contact.HomeAddressState = value; }

        public string HomeAddressStreet { get => contact.HomeAddressStreet; set => contact.HomeAddressStreet = value; }

        public string HomeFaxNumber { get => contact.HomeFaxNumber; set => contact.HomeFaxNumber = value; }

        public string HomeTelephoneNumber { get => contact.HomeTelephoneNumber; set => contact.HomeTelephoneNumber = value; }

        public string Initials { get => contact.Initials; set => contact.Initials = value; }

        public string InternetFreeBusyAddress { get => contact.InternetFreeBusyAddress; set => contact.InternetFreeBusyAddress = value; }

        public string ISDNNumber { get => contact.ISDNNumber; set => contact.ISDNNumber = value; }

        public string JobTitle { get => contact.JobTitle; set => contact.JobTitle = value; }

        public bool Journal { get => contact.Journal; set => contact.Journal = value; }

        public string Language { get => contact.Language; set => contact.Language = value; }

        public string LastFirstAndSuffix => contact.LastFirstAndSuffix;

        public string LastFirstNoSpace => contact.LastFirstNoSpace;

        public string LastFirstNoSpaceCompany => contact.LastFirstNoSpaceCompany;

        public string LastFirstSpaceOnly => contact.LastFirstSpaceOnly;

        public string LastFirstSpaceOnlyCompany => contact.LastFirstSpaceOnlyCompany;

        public string LastName { get => contact.LastName; set => contact.LastName = value; }

        public string LastNameAndFirstName => contact.LastNameAndFirstName;

        public string MailingAddress { get => contact.MailingAddress; set => contact.MailingAddress = value; }

        public string MailingAddressCity { get => contact.MailingAddressCity; set => contact.MailingAddressCity = value; }

        public string MailingAddressCountry { get => contact.MailingAddressCountry; set => contact.MailingAddressCountry = value; }

        public string MailingAddressPostalCode { get => contact.MailingAddressPostalCode; set => contact.MailingAddressPostalCode = value; }

        public string MailingAddressPostOfficeBox { get => contact.MailingAddressPostOfficeBox; set => contact.MailingAddressPostOfficeBox = value; }

        public string MailingAddressState { get => contact.MailingAddressState; set => contact.MailingAddressState = value; }

        public string MailingAddressStreet { get => contact.MailingAddressStreet; set => contact.MailingAddressStreet = value; }

        public string ManagerName { get => contact.ManagerName; set => contact.ManagerName = value; }

        public string MiddleName { get => contact.MiddleName; set => contact.MiddleName = value; }

        public string MobileTelephoneNumber { get => contact.MobileTelephoneNumber; set => contact.MobileTelephoneNumber = value; }

        public string NetMeetingAlias { get => contact.NetMeetingAlias; set => contact.NetMeetingAlias = value; }

        public string NetMeetingServer { get => contact.NetMeetingServer; set => contact.NetMeetingServer = value; }

        public string NickName { get => contact.NickName; set => contact.NickName = value; }

        public string OfficeLocation { get => contact.OfficeLocation; set => contact.OfficeLocation = value; }

        public string OrganizationalIDNumber { get => contact.OrganizationalIDNumber; set => contact.OrganizationalIDNumber = value; }

        public string OtherAddress { get => contact.OtherAddress; set => contact.OtherAddress = value; }

        public string OtherAddressCity { get => contact.OtherAddressCity; set => contact.OtherAddressCity = value; }

        public string OtherAddressCountry { get => contact.OtherAddressCountry; set => contact.OtherAddressCountry = value; }

        public string OtherAddressPostalCode { get => contact.OtherAddressPostalCode; set => contact.OtherAddressPostalCode = value; }

        public string OtherAddressPostOfficeBox { get => contact.OtherAddressPostOfficeBox; set => contact.OtherAddressPostOfficeBox = value; }

        public string OtherAddressState { get => contact.OtherAddressState; set => contact.OtherAddressState = value; }

        public string OtherAddressStreet { get => contact.OtherAddressStreet; set => contact.OtherAddressStreet = value; }

        public string OtherFaxNumber { get => contact.OtherFaxNumber; set => contact.OtherFaxNumber = value; }

        public string OtherTelephoneNumber { get => contact.OtherTelephoneNumber; set => contact.OtherTelephoneNumber = value; }

        public string PagerNumber { get => contact.PagerNumber; set => contact.PagerNumber = value; }

        public string PersonalHomePage { get => contact.PersonalHomePage; set => contact.PersonalHomePage = value; }

        public string PrimaryTelephoneNumber { get => contact.PrimaryTelephoneNumber; set => contact.PrimaryTelephoneNumber = value; }

        public string Profession { get => contact.Profession; set => contact.Profession = value; }

        public string RadioTelephoneNumber { get => contact.RadioTelephoneNumber; set => contact.RadioTelephoneNumber = value; }

        public string ReferredBy { get => contact.ReferredBy; set => contact.ReferredBy = value; }

        public OlMailingAddress SelectedMailingAddress { get => contact.SelectedMailingAddress; set => contact.SelectedMailingAddress = value; }

        public string Spouse { get => contact.Spouse; set => contact.Spouse = value; }

        public string Suffix { get => contact.Suffix; set => contact.Suffix = value; }

        public string TelexNumber { get => contact.TelexNumber; set => contact.TelexNumber = value; }

        public string Title { get => contact.Title; set => contact.Title = value; }

        public string TTYTDDTelephoneNumber { get => contact.TTYTDDTelephoneNumber; set => contact.TTYTDDTelephoneNumber = value; }

        public string User1 { get => contact.User1; set => contact.User1 = value; }

        public string User2 { get => contact.User2; set => contact.User2 = value; }

        public string User3 { get => contact.User3; set => contact.User3 = value; }

        public string User4 { get => contact.User4; set => contact.User4 = value; }

        public string UserCertificate { get => contact.UserCertificate; set => contact.UserCertificate = value; }

        public string WebPage { get => contact.WebPage; set => contact.WebPage = value; }

        public string YomiCompanyName { get => contact.YomiCompanyName; set => contact.YomiCompanyName = value; }

        public string YomiFirstName { get => contact.YomiFirstName; set => contact.YomiFirstName = value; }

        public string YomiLastName { get => contact.YomiLastName; set => contact.YomiLastName = value; }

        public Links Links => contact.Links;

        public ItemProperties ItemProperties => contact.ItemProperties;

        public string LastFirstNoSpaceAndSuffix => contact.LastFirstNoSpaceAndSuffix;

        public OlDownloadState DownloadState => contact.DownloadState;

        public string IMAddress { get => contact.IMAddress; set => contact.IMAddress = value; }

        public OlRemoteStatus MarkForDownload { get => contact.MarkForDownload; set => contact.MarkForDownload = value; }

        public bool IsConflict => contact.IsConflict;

        public bool AutoResolvedWinner => contact.AutoResolvedWinner;

        public Conflicts Conflicts => contact.Conflicts;

        public bool HasPicture => contact.HasPicture;

        public PropertyAccessor PropertyAccessor => contact.PropertyAccessor;

        public string TaskSubject { get => contact.TaskSubject; set => contact.TaskSubject = value; }

        public DateTime TaskDueDate { get => contact.TaskDueDate; set => contact.TaskDueDate = value; }

        public DateTime TaskStartDate { get => contact.TaskStartDate; set => contact.TaskStartDate = value; }

        public DateTime TaskCompletedDate { get => contact.TaskCompletedDate; set => contact.TaskCompletedDate = value; }

        public DateTime ToDoTaskOrdinal { get => contact.ToDoTaskOrdinal; set => contact.ToDoTaskOrdinal = value; }

        public bool ReminderOverrideDefault { get => contact.ReminderOverrideDefault; set => contact.ReminderOverrideDefault = value; }

        public bool ReminderPlaySound { get => contact.ReminderPlaySound; set => contact.ReminderPlaySound = value; }

        public bool ReminderSet { get => contact.ReminderSet; set => contact.ReminderSet = value; }

        public string ReminderSoundFile { get => contact.ReminderSoundFile; set => contact.ReminderSoundFile = value; }

        public DateTime ReminderTime { get => contact.ReminderTime; set => contact.ReminderTime = value; }

        public bool IsMarkedAsTask => contact.IsMarkedAsTask;

        public string BusinessCardLayoutXml { get => contact.BusinessCardLayoutXml; set => contact.BusinessCardLayoutXml = value; }

        public OlBusinessCardType BusinessCardType => contact.BusinessCardType;

        public string ConversationID => contact.ConversationID;

        public dynamic RTFBody { get => contact.RTFBody; set => contact.RTFBody = value; }

        public event ItemEvents_10_OpenEventHandler Open
        {
            add
            {
                contact.Open += value;
            }

            remove
            {
                contact.Open -= value;
            }
        }

        public event ItemEvents_10_CustomActionEventHandler CustomAction
        {
            add
            {
                contact.CustomAction += value;
            }

            remove
            {
                contact.CustomAction -= value;
            }
        }

        public event ItemEvents_10_CustomPropertyChangeEventHandler CustomPropertyChange
        {
            add
            {
                contact.CustomPropertyChange += value;
            }

            remove
            {
                contact.CustomPropertyChange -= value;
            }
        }

        public event ItemEvents_10_ForwardEventHandler Forward
        {
            add
            {
                contact.Forward += value;
            }

            remove
            {
                contact.Forward -= value;
            }
        }

        event ItemEvents_10_CloseEventHandler ItemEvents_10_Event.Close
        {
            add
            {
                ((ItemEvents_10_Event)contact).Close += value;
            }

            remove
            {
                ((ItemEvents_10_Event)contact).Close -= value;
            }
        }

        public event ItemEvents_10_PropertyChangeEventHandler PropertyChange
        {
            add
            {
                contact.PropertyChange += value;
            }

            remove
            {
                contact.PropertyChange -= value;
            }
        }

        public event ItemEvents_10_ReadEventHandler Read
        {
            add
            {
                contact.Read += value;
            }

            remove
            {
                contact.Read -= value;
            }
        }

        public event ItemEvents_10_ReplyEventHandler Reply
        {
            add
            {
                contact.Reply += value;
            }

            remove
            {
                contact.Reply -= value;
            }
        }

        public event ItemEvents_10_ReplyAllEventHandler ReplyAll
        {
            add
            {
                contact.ReplyAll += value;
            }

            remove
            {
                contact.ReplyAll -= value;
            }
        }

        public event ItemEvents_10_SendEventHandler Send
        {
            add
            {
                contact.Send += value;
            }

            remove
            {
                contact.Send -= value;
            }
        }

        public event ItemEvents_10_WriteEventHandler Write
        {
            add
            {
                contact.Write += value;
            }

            remove
            {
                contact.Write -= value;
            }
        }

        public event ItemEvents_10_BeforeCheckNamesEventHandler BeforeCheckNames
        {
            add
            {
                contact.BeforeCheckNames += value;
            }

            remove
            {
                contact.BeforeCheckNames -= value;
            }
        }

        public event ItemEvents_10_AttachmentAddEventHandler AttachmentAdd
        {
            add
            {
                contact.AttachmentAdd += value;
            }

            remove
            {
                contact.AttachmentAdd -= value;
            }
        }

        public event ItemEvents_10_AttachmentReadEventHandler AttachmentRead
        {
            add
            {
                contact.AttachmentRead += value;
            }

            remove
            {
                contact.AttachmentRead -= value;
            }
        }

        public event ItemEvents_10_BeforeAttachmentSaveEventHandler BeforeAttachmentSave
        {
            add
            {
                contact.BeforeAttachmentSave += value;
            }

            remove
            {
                contact.BeforeAttachmentSave -= value;
            }
        }

        public event ItemEvents_10_BeforeDeleteEventHandler BeforeDelete
        {
            add
            {
                contact.BeforeDelete += value;
            }

            remove
            {
                contact.BeforeDelete -= value;
            }
        }

        public event ItemEvents_10_AttachmentRemoveEventHandler AttachmentRemove
        {
            add
            {
                contact.AttachmentRemove += value;
            }

            remove
            {
                contact.AttachmentRemove -= value;
            }
        }

        public event ItemEvents_10_BeforeAttachmentAddEventHandler BeforeAttachmentAdd
        {
            add
            {
                contact.BeforeAttachmentAdd += value;
            }

            remove
            {
                contact.BeforeAttachmentAdd -= value;
            }
        }

        public event ItemEvents_10_BeforeAttachmentPreviewEventHandler BeforeAttachmentPreview
        {
            add
            {
                contact.BeforeAttachmentPreview += value;
            }

            remove
            {
                contact.BeforeAttachmentPreview -= value;
            }
        }

        public event ItemEvents_10_BeforeAttachmentReadEventHandler BeforeAttachmentRead
        {
            add
            {
                contact.BeforeAttachmentRead += value;
            }

            remove
            {
                contact.BeforeAttachmentRead -= value;
            }
        }

        public event ItemEvents_10_BeforeAttachmentWriteToTempFileEventHandler BeforeAttachmentWriteToTempFile
        {
            add
            {
                contact.BeforeAttachmentWriteToTempFile += value;
            }

            remove
            {
                contact.BeforeAttachmentWriteToTempFile -= value;
            }
        }

        public event ItemEvents_10_UnloadEventHandler Unload
        {
            add
            {
                contact.Unload += value;
            }

            remove
            {
                contact.Unload -= value;
            }
        }

        public event ItemEvents_10_BeforeAutoSaveEventHandler BeforeAutoSave
        {
            add
            {
                contact.BeforeAutoSave += value;
            }

            remove
            {
                contact.BeforeAutoSave -= value;
            }
        }

        public event ItemEvents_10_BeforeReadEventHandler BeforeRead
        {
            add
            {
                contact.BeforeRead += value;
            }

            remove
            {
                contact.BeforeRead -= value;
            }
        }

        public event ItemEvents_10_AfterWriteEventHandler AfterWrite
        {
            add
            {
                contact.AfterWrite += value;
            }

            remove
            {
                contact.AfterWrite -= value;
            }
        }

        public event ItemEvents_10_ReadCompleteEventHandler ReadComplete
        {
            add
            {
                contact.ReadComplete += value;
            }

            remove
            {
                contact.ReadComplete -= value;
            }
        }

        #endregion

        #region Properties

        private string inboxPath;
        /// <summary>
        /// A Path to an Inbox folder where the emails for this contact are stored.
        /// </summary>
        public string InboxPath
        {
            get { return inboxPath; }
            set { inboxPath = value; }
        }

        private string sentPath;
        /// <summary>
        /// A Path to an SentItems folder.
        /// 
        /// where the emails for this contact are stored.
        /// </summary>
        public string SentPath
        {
            get { return sentPath; }
            set { sentPath = value; }
        }

        private EmailAction deliveryAction = EmailAction.None;
        public EmailAction DeliveryAction
        {
            get { return deliveryAction; }
            set { deliveryAction = value; }
        }

        private EmailAction readAction = EmailAction.None;
        public EmailAction ReadAction
        {
            get { return readAction; }
            set { readAction = value; }
        }

        private EmailAction sentAction = EmailAction.None;
        public EmailAction SentAction
        {
            get { return sentAction; }
            set { sentAction = value; }
        }

        private bool samePath = true;

        public bool SamePath
        {
            get { return samePath; }
            set { samePath = value; }
        }

        #endregion

        public InTouchContact(ContactItem contact) : base()
        {
            this.contact = contact;
            GetData();
        }

        public void SaveAndDispose()
        {
            SaveData();
            if (contact is object) Marshal.ReleaseComObject(contact);
        }

        #region Data Methods

        private void GetData()
        {
            string data;

            if (contact is object)
            {
                UserProperty customProperty = contact.UserProperties.Find("InTouchContact" );
                if (customProperty is object)
                {
                    data = customProperty.Value;
                }
                else
                {
                    contact.UserProperties.Add("InTouchContact", OlUserPropertyType.olText);
                    data = "|";
                    data += "|";
                    data += "0|";
                    data += "0|";
                    data += "0|";
                    data += "True|";

                    contact.UserProperties["InTouchContact"].Value = data;
                    contact.Save();
                }
                if (customProperty is object) Marshal.ReleaseComObject(customProperty);
            }
            else
            {
                data = "|";
                data += "|";
                data += "0|";
                data += "0|";
                data += "0|";
                data += "True|";
            }

            string[] values = data.Split('|');            

            try
            {
                inboxPath = values[0];
                sentPath = values[1];
                deliveryAction = (EmailAction)int.Parse(values[2]);
                readAction = (EmailAction)int.Parse(values[3]);
                sentAction = (EmailAction)int.Parse(values[4]);
                samePath = bool.Parse(values[5]);
            }
            catch(FormatException)
            {
                //Contact data is from an older format. It will be saved in the newer format when the contact is saved.
                Op.LogMessage("Exception Managed.");
            }
            catch(System.Exception ex)
            {
                Op.LogError(ex);
                throw;
            }
        }

        private void SaveData()
        {
            string data;
            data = inboxPath + "|";
            data += sentPath + "|";
            data += (int)deliveryAction + "|";
            data += (int)readAction + "|";
            data += (int)sentAction + "|";
            data += samePath.ToString() + "|";

            UserProperty customProperty = contact.UserProperties.Find("InTouchContact");
            if (customProperty is object)
            {
                customProperty.Value = data;
            }
            else
            {
                contact.UserProperties.Add("InTouchContact", OlUserPropertyType.olText);
                contact.UserProperties["InTouchContact"].Value = data;
            }
            contact.Save();
            if (customProperty is object) Marshal.ReleaseComObject(customProperty);
        }

        #endregion

        public string GetContactPicturePath()
        {
            string picturePath = "";
            if (contact.HasPicture)
            {
                foreach (Attachment nextAttachment in contact.Attachments)
                {
                    if (nextAttachment.DisplayName == "ContactPicture.jpg")
                    {
                        try
                        {
                            picturePath = System.IO.Path.GetDirectoryName(System.IO.Path.GetTempPath()) + "\\Contact_" + contact.EntryID + ".jpg";
                            if (!System.IO.File.Exists(picturePath))
                            {
                                nextAttachment.SaveAsFile(picturePath);
                            }
                        }
                        catch (COMException)
                        {
                            picturePath = "";
                        }
                        catch (System.Exception ex)
                        {
                            Op.LogError(ex);
                            picturePath = "";
                            throw;
                        }
                    }
                }
            }
            return picturePath;
        }

        /// <summary>
        /// Check if the extention properties are valid for this contact.
        /// </summary>
        /// <returns>Returns false is there are any problems with the properties.</returns>
        public bool CheckDetails()
        {
            bool returnValue = true;

            if (inboxPath is object)
            {
                //TODO Replace Op with the Contact object.
                if (!Op.CheckFolderPath(inboxPath))
                {
                    returnValue = false;
                    Op.LogMessage(FullName + " Inbox Folder Path is Invalid : " + inboxPath);
                }
            }
            else
            {
                returnValue = false;
            }

            if(sentPath is object)
            {
                if (!Op.CheckFolderPath(sentPath))
                {
                    returnValue = false;
                    Op.LogMessage(FullName + " Sent Folder Path is Invalid : " + sentPath);
                }
            }
            else
            {
                returnValue = false;
            }

            return returnValue;
        }
    }
}