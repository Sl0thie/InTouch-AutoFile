using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InTouch_AutoFile
{
    internal class TaskFileInbox
    {
        private readonly Action callBack;
        private readonly IList<Outlook.MailItem> mailToProcess = new List<Outlook.MailItem>();

        public TaskFileInbox(Action callBack)
        {
            this.callBack = callBack;
        }

        public void RunTask()
        {
            //If task is enabled in the settings then start task.
            if (Properties.Settings.Default.TaskInbox)
            {
                Op.LogMessage("Starting FileInbox Task.");
                Thread backgroundThread = new Thread(new ThreadStart(BackgroundProcess))
                {
                    Name = "AF.FileInbox",
                    IsBackground = true,
                    Priority = ThreadPriority.Normal
                };
                backgroundThread.SetApartmentState(ApartmentState.STA);
                backgroundThread.Start();
            }
        }

        private void BackgroundProcess()
        {
            CreateListOfInboxItems();
            ProcessListOfItems();
            callBack?.Invoke();
        }

        /// <summary>
        /// Create a List of items within the Inbox. Exclude appointments as well as flagged emails.
        /// </summary>
        private void CreateListOfInboxItems()
        {
            foreach (object nextItem in Globals.ThisAddIn.Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items)
            {
                if (nextItem is Outlook.MailItem email)
                {
                    //Only process emails that don't have a flag.
                    switch (email.FlagRequest)
                    {
                        case "":
                            mailToProcess.Add(email);
                            break;
                        case "Follow up":
                            //Don't process follow up. This leave them in the inbox for manual processing.
                            Op.LogMessage("Move Email : Email has a flag set.");
                            break;
                        case null:
                            mailToProcess.Add(email);
                            break;
                        default:
                            Op.LogMessage("Move Email : Unknown Flag Request Type.");
                            break;
                    }
                }
            }
        }

        private void ProcessListOfItems()
        {
            foreach (Outlook.MailItem nextEmail in mailToProcess)
            {
                ProcessEmail(nextEmail);
            }
        }

        private static void ProcessEmail(Outlook.MailItem email)
        {
            //Email may have been deleted or moved so check if it exists first.
            if (email is object)
            {
                //Check if the email has a Sender.
                if (email.Sender is object)
                {
                    //Find the Contact accociated with the Sender.
                    InTouchContact mailContact = null;
                    Outlook.ContactItem contact = InTouch.Contacts.FindContactFromEmailAddress(email.Sender.Address);
                    if (contact is object)
                    {
                        mailContact = new InTouchContact(contact);
                    }

                    //If found then try to process the email.
                    if (mailContact is object)
                    {
                        //If unread the process delivery option else process read option.
                        if (email.UnRead)
                        {
                            switch (mailContact.DeliveryAction)
                            {
                                case EmailAction.None: //Don't do anything to the email.
                                    Op.LogMessage("Move Email : Delivery Action set to None. " + email.Sender.Address);
                                    break;

                                case EmailAction.Delete: //Delete the email if it is passed its action date.
                                    Op.LogMessage("Move Email : Deleting email from " + email.Sender.Address);
                                    email.Delete();
                                    break;

                                case EmailAction.Move: //Move the email if its passed its action date.
                                    Op.LogMessage("Move Email : Moving email from " + email.Sender.Address);
                                    MoveEmailToFolder(mailContact.InboxPath, email);
                                    break;
                            }
                        }
                        else
                        {
                            switch (mailContact.ReadAction)
                            {
                                case EmailAction.None: //Don't do anything to the email.
                                    Op.LogMessage("Move Email : Read Action set to None. " + email.Sender.Address);
                                    break;

                                case EmailAction.Delete: //Delete the email.
                                    Op.LogMessage("Move Email : Deleting email from " + email.Sender.Address);
                                    email.Delete();
                                    break;

                                case EmailAction.Move: //Move the email.
                                    Op.LogMessage("Move Email : Moving email from " + email.Sender.Address);
                                    MoveEmailToFolder(mailContact.InboxPath, email);
                                    break;
                            }
                        }
                        mailContact.SaveAndDispose();
                    }
                    else //If not found then just log it for the moment.
                    {
                        try
                        {
                            //Get the 'On Behalf' property from the email.
                            Outlook.PropertyAccessor mapiPropertyAccessor;
                            string propertyName = "http://schemas.microsoft.com/mapi/proptag/0x0065001F";
                            mapiPropertyAccessor = email.PropertyAccessor;
                            string onBehalfEmailAddress = mapiPropertyAccessor.GetProperty(propertyName).ToString();
                            if (mapiPropertyAccessor is object)
                            {
                                Marshal.ReleaseComObject(mapiPropertyAccessor);
                            }

                            //Log the details.                           
                            Op.LogMessage("Move Email : No Contact for " + email.SenderEmailAddress);
                            Op.LogMessage("SenderName         : " + email.SenderName);
                            Op.LogMessage("SentOnBehalfOfName : " + email.SentOnBehalfOfName);
                            Op.LogMessage("ReplyRecipientNames: " + email.ReplyRecipientNames);
                            Op.LogMessage("On Behalf: " + onBehalfEmailAddress);
                            Op.LogMessage("");
                        }
                        catch (Exception ex)
                        {
                            Op.LogError(ex);
                            throw;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Method to move the email from the Inbox to the specified folder.
        /// </summary>
        /// <param name="folderPath">The path to the folder to move the email.</param>
        /// <param name="email">The mailitem to move.</param>
        private static void MoveEmailToFolder(string folderPath, Outlook.MailItem email)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
            string[] folders = folderPath.Split('\\');

            for (int i = 0; i <= folders.GetUpperBound(0); i++)
            {
                Outlook.Folders subFolders = folder.Folders;
                folder = subFolders[folders[i]] as Outlook.Folder;
            }

            email.Move(folder);
        }
    }
}