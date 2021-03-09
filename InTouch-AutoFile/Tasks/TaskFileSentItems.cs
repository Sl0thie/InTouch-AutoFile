using System;
using System.Collections.Generic;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InTouch_AutoFile
{
    internal class TaskFileSentItems
    {
        private readonly Action callBack;
        private readonly IList<Outlook.MailItem> mailToProcess = new List<Outlook.MailItem>();

        public TaskFileSentItems(Action callBack)
        {
            this.callBack = callBack;
        }

        public void RunTask()
        {
            Thread backgroundThread = new Thread(new ThreadStart(BackgroundProcess))
            {
                Name = "InTouch_Backend.TaskFileSentItems",
                IsBackground = true,
                Priority = ThreadPriority.Normal
            };
            backgroundThread.SetApartmentState(ApartmentState.STA);
            backgroundThread.Start();
        }

        private void BackgroundProcess()
        {
            try
            {
                Op.LogMessage("Starting TaskFileSentItems Task.");

                //foreach (Outlook.Folder nextfolder in contactsFolder.Folders)
                //{
                //    if (nextfolder.Name == "Other Contacts")
                //    {
                //        ContactFolders.Add(nextfolder.Name);
                //        Op.LogMessage("Adding Contacts Folder : " + nextfolder.Name);
                //    }
                //    if (nextfolder.Name == "Unknown Contacts")
                //    {
                //        ContactFolders.Add(nextfolder.Name);
                //        Op.LogMessage("Adding Contacts Folder : " + nextfolder.Name);
                //    }
                //    if (nextfolder.Name == "Alias")
                //    {
                //        ContactFolders.Add(nextfolder.Name);
                //        Op.LogMessage("Adding Contacts Folder : " + nextfolder.Name);
                //    }
                //}
                CreateListOfInboxItems();
                ProcessListOfITems();
            }
            catch (Exception ex) { Op.LogError(ex); throw; }
            callBack?.Invoke();
        }

        private void ProcessListOfITems()
        {
            foreach (Outlook.MailItem nextItem in mailToProcess)
            {
                CheckEmail(nextItem);
            }
        }

        private void CreateListOfInboxItems()
        {
            foreach (object mailItem in Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items)
            {
                try
                {
                    if (mailItem is Outlook.MailItem item)
                    {
                        mailToProcess.Add(item);
                    }
                }
                catch (Exception ex) { Op.LogError(ex); throw; }
            }
        }

        private static void CheckEmail(Outlook.MailItem Email)
        {
            //if (!ReferenceEquals(null, Email)) // recently sent items can move sometimes.
            //{
            //    Outlook.Recipients recipients = Email.Recipients;
            //    Outlook.Recipient recipient = recipients[1];
            //    if (!ReferenceEquals(null, recipient))
            //    {

            //        Outlook.ContactItem contact = InTouch.FindContactFromEmailAddress(recipient.Address);
            //        if (!object.ReferenceEquals(contact, null))
            //        {
            //            //Op.LogMessage("Found Contact for : " + recipient.Address);
            //            ProcessEmailWithContact(Email, contact);
            //        }
            //        else
            //        {
            //            Op.LogMessage("No Contact for : " + recipient.Address);
            //        }
            //        if (contact != null) Marshal.ReleaseComObject(contact);
            //    }
            //    else
            //    {
            //        Op.LogMessage("No Recipient : ");
            //    }
            //}
        }

        //private Outlook.Items GetAppointmentsWithSpawn(string SpawnGUID)
        //{
        //    Outlook.Folder emailsCalendar = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Folders["Emails"] as Outlook.Folder;


        //    string filter = "[Subject] = '"
        //        + SpawnGUID
        //        + "'";
        //    //Debug.WriteLine(filter);
        //    try
        //    {
        //        Outlook.Items calItems = emailsCalendar.Items;
        //        calItems.IncludeRecurrences = true;
        //        calItems.Sort("[Start]", Type.Missing);
        //        Outlook.Items restrictItems = calItems.Restrict(filter);
        //        if (restrictItems.Count > 0)
        //        {
        //            return restrictItems;
        //        }
        //        else
        //        {
        //            return null;
        //        }
        //    }
        //    catch { return null; }
        //}

        //private void ProcessEmailWithContact(Outlook.MailItem mailItem, Outlook.ContactItem contactItem)
        //{
        //    //ExContact exContact = new ExContact(contactItem);
        //    //ExMailItem exMailItem = new ExMailItem(mailItem);

        //    //if (!exMailItem.AddedToContact)
        //    //{
        //    //    exMailItem.ContactEntryId = exContact.EntryId;
        //    //    if (!string.IsNullOrWhiteSpace(exContact.ObjectiveId))
        //    //    {
        //    //        string[] temp = exContact.ObjectiveId.Split('|');
        //    //        exMailItem.ObjectiveId = temp[0];
        //    //    }

        //    //    exContact.AddConnection(exMailItem.SentOn, ConnectionType.Email, ConnectionDirection.Outgoing);

        //    //    exMailItem.AddedToContact = true;

        //    //    Outlook.AppointmentItem NewAppointmentItem = null;
        //    //    NewAppointmentItem = (Outlook.AppointmentItem)EmailCalendarFolder.Items.Add(Outlook.OlItemType.olAppointmentItem);

        //    //    NewAppointmentItem.Subject = "Sent Email : " + exMailItem.Subject;
        //    //    NewAppointmentItem.Categories = "Email - Sent";
        //    //    NewAppointmentItem.Body = exMailItem.EntryId + "|" + exContact.EntryId;
        //    //    NewAppointmentItem.ReminderSet = false;

        //    //    NewAppointmentItem.Start = Op.OutlookDateTime(exMailItem.SentOn);
        //    //    NewAppointmentItem.End = Op.OutlookDateTime(exMailItem.SentOn.AddMinutes(1));

        //    //    NewAppointmentItem.Save();



        //    //}

        //    //exMailItem.Save();
        //    //exContact.Save();

        //    //string SpawnGUID = "";
        //    //Outlook.UserProperty CustomProperty = mailItem.UserProperties.Find("SpawnGUID");
        //    //if (!object.ReferenceEquals(CustomProperty, null))
        //    //{
        //    //    SpawnGUID = CustomProperty.Value;

        //    //}
        //    //if (CustomProperty != null) Marshal.ReleaseComObject(CustomProperty);

        //    //if (!string.IsNullOrWhiteSpace(SpawnGUID))
        //    //{
        //    //    if (SpawnGUID.Substring(0, 6) == "Spawn ")
        //    //    {
        //    //        Outlook.Items spawnItems = GetAppointmentsWithSpawn(SpawnGUID);

        //    //        foreach (Outlook.AppointmentItem nextItem in spawnItems)
        //    //        {
        //    //            nextItem.Subject = "Write Email : " + exMailItem.Subject;
        //    //            nextItem.Categories = "Email - Write";
        //    //            nextItem.Body = exMailItem.EntryId + "|" + exContact.EntryId;
        //    //            nextItem.ReminderSet = false;
        //    //            nextItem.Save();

        //    //            exContact.AddTime(nextItem.Start, nextItem.End, ConnectionType.Email, ConnectionDirection.Outgoing);

        //    //            string bodyString = exMailItem.EntryId + "|";

        //    //            ExObjectiveDetails exObjectiveDetails = new ExObjectiveDetails(nextItem);
        //    //            exObjectiveDetails.ItemEntryId = exMailItem.EntryId;
        //    //            exObjectiveDetails.ItemSubject = mailItem.Subject;
        //    //            exObjectiveDetails.ItemType = ItemType.EmailOut;

        //    //            if (!ReferenceEquals(null, contactItem))
        //    //            {
        //    //                bodyString += contactItem.EntryID + "|";
        //    //                exObjectiveDetails.ContactEntryId = contactItem.EntryID;

        //    //                string temp = contactItem.User1;
        //    //                if (!ReferenceEquals(null, temp))
        //    //                {
        //    //                    if (temp.IndexOf('|') >= 0)
        //    //                    {
        //    //                        string[] temps = temp.Split('|');
        //    //                        exObjectiveDetails.ObjectiveGuid = Guid.Parse(temps[0]);
        //    //                        exObjectiveDetails.ObjectiveName = temps[1];
        //    //                        Op.LogMessage("Task Sent Items Parse GUID " + exObjectiveDetails.ObjectiveGuid.ToString());

        //    //                        exObjectiveDetails.ObjectiveName = temps[1];
        //    //                    }
        //    //                }
        //    //            }
        //    //            exObjectiveDetails.Save();

        //    //            nextItem.Body = bodyString;
        //    //            nextItem.Save();

        //    //        }
        //    //    }
        //    //}

        //    //if (!string.IsNullOrWhiteSpace(exContact.AutoSortSent))
        //    //{
        //    //    MoveMail(exContact.AutoSortSent, mailItem);
        //    //}
        //    //else
        //    //{
        //    //    Op.LogMessage("No AutoSortSent for " + exContact.FullName);
        //    //}

        //    //exMailItem = null;
        //    //exContact = null;
        //    //if (contactItem != null) Marshal.ReleaseComObject(contactItem);
        //    //if (mailItem != null) Marshal.ReleaseComObject(mailItem);
        //}

        //private void MoveMail(string FolderString, Outlook.MailItem MailItem)
        //{
        //    try
        //    {
        //        Op.LogMessage("Moving Sent Mail " + MailItem.SenderEmailAddress.ToLower());
        //        string[] FoldersString = FolderString.Split('\\');
        //        switch (FoldersString.Count())
        //        {
        //            case 1:
        //                MailItem.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Folders[FoldersString[0]]);
        //                break;
        //            case 2:
        //                MailItem.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Folders[FoldersString[0]].Folders[FoldersString[1]]);
        //                break;
        //            case 3:
        //                MailItem.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Folders[FoldersString[0]].Folders[FoldersString[1]].Folders[FoldersString[2]]);
        //                break;
        //            case 4:
        //                MailItem.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Folders[FoldersString[0]].Folders[FoldersString[1]].Folders[FoldersString[2]].Folders[FoldersString[3]]);
        //                break;
        //            case 5:
        //                MailItem.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Folders[FoldersString[0]].Folders[FoldersString[1]].Folders[FoldersString[2]].Folders[FoldersString[3]].Folders[FoldersString[4]]);
        //                break;
        //            case 6:
        //                MailItem.Move(Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Folders[FoldersString[0]].Folders[FoldersString[1]].Folders[FoldersString[2]].Folders[FoldersString[3]].Folders[FoldersString[4]].Folders[FoldersString[5]]);
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ex.Message != "The attempted operation failed.  An object could not be found.")
        //        {
        //            Op.LogError(ex);
        //        }
        //    }
        //}

        //private bool FindContactDetailsFromEmailAddress(string EmailAddress)
        //{
        //    EmailAddress = EmailAddress.ToLower();

        //    try
        //    {
        //        if (SearchFolderForContactFromEmailAddress("", EmailAddress))
        //        {
        //            return true;
        //        }
        //        foreach (string NextContactFolder in ContactFolders)
        //        {
        //            if (SearchFolderForContactFromEmailAddress(NextContactFolder, EmailAddress))
        //            {
        //                return true;
        //            }
        //        }

        //        //If not found then return false to mark as unknown.
        //        ContactItem = null;
        //        return false;
        //    }
        //    catch (Exception ex) { Op.LogError(ex); return false; }
        //}

        //private bool SearchFolderForContactFromEmailAddress(string FolderName, string EmailAddress)
        //{
        //    bool returnValue = false;
        //    try
        //    {
        //        Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
        //        Outlook.MAPIFolder contactsFolder;
        //        if (FolderName == "")
        //        {
        //            contactsFolder = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
        //        }
        //        else
        //        {
        //            contactsFolder = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts).Folders[FolderName] as Outlook.Folder;
        //        }

        //        foreach (var NextObject in contactsFolder.Items)
        //        {
        //            if (NextObject is Outlook.ContactItem)
        //            {
        //                Outlook.ContactItem NextContact = (Outlook.ContactItem)NextObject;

        //                if (!object.ReferenceEquals(null, NextContact.Email1Address))
        //                {
        //                    if (NextContact.Email1Address.ToLower() == EmailAddress)
        //                    {
        //                        ContactItem = NextContact;
        //                        returnValue = true;
        //                        break;
        //                    }
        //                }

        //                if (!object.ReferenceEquals(null, NextContact.Email2Address))
        //                {
        //                    if (NextContact.Email2Address.ToLower() == EmailAddress)
        //                    {
        //                        ContactItem = NextContact;
        //                        returnValue = true;
        //                        break;
        //                    }
        //                }

        //                if (!object.ReferenceEquals(null, NextContact.Email3Address))
        //                {
        //                    if (NextContact.Email3Address.ToLower() == EmailAddress)
        //                    {
        //                        ContactItem = NextContact;
        //                        returnValue = true;
        //                        break;
        //                    }
        //                }

        //                if (NextContact != null) Marshal.ReleaseComObject(NextContact);
        //            }
        //            if (NextObject != null) Marshal.ReleaseComObject(NextObject);
        //        }

        //        if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
        //        if (outlookNameSpace != null) Marshal.ReleaseComObject(outlookNameSpace);
        //        return returnValue;
        //    }
        //    catch (Exception ex) { Op.LogError(ex); return false; }
        //}
    }
}