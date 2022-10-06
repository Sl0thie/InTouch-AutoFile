namespace InTouch_AutoFile
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using Serilog;

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
                Log.Information("Starting FileInbox Task.");
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
                            Log.Information("Move Email : Email has a flag set.");
                            break;
                        case null:
                            mailToProcess.Add(email);
                            break;
                        default:
                            Log.Information("Move Email : Unknown Flag Request Type.");
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
            bool ok = true;

            try
            {
                if(email.Sender is null)
                {
                    ok = false;
                }
            }
            catch(Exception ex)
            {
                Log.Error(ex.Message, ex);
                ok = false;
            }

            InTouchContact mailContact = null;
            Outlook.ContactItem contact = null ;

            try
            {
                contact = InTouch.Contacts.FindContactFromEmailAddress(email.Sender.Address);
            }
            catch (InvalidComObjectException ex)
            {
                Log.Error(ex.Message, ex);
            }

            if (contact is object)
            {
                mailContact = new InTouchContact(contact);
            }
            else
            {
                ok = false;
            }
            if (mailContact is null) { ok = false; }

            if (ok)
            {
                //If unread the process delivery option else process read option.
                if (email.UnRead)
                {
                    switch (mailContact.DeliveryAction)
                    {
                        case EmailAction.None: //Don't do anything to the email.
                            Log.Information("Move Email : Delivery Action set to None. " + email.Sender.Address);
                            break;

                        case EmailAction.Delete: //Delete the email if it is passed its action date.
                            Log.Information("Move Email : Deleting email from " + email.Sender.Address);
                            email.Delete();
                            break;

                        case EmailAction.Move: //Move the email if its passed its action date.
                            Log.Information("Move Email : Moving email from " + email.Sender.Address);
                            MoveEmailToFolder(mailContact.InboxPath, email);
                            break;

                        case EmailAction.Junk:
                            Log.Information("Move Email to Junk: Moving email from " + email.Sender.Address);
                            email.Move(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk));
                            break;
                    }
                }
                else
                {
                    switch (mailContact.ReadAction)
                    {
                        case EmailAction.None: //Don't do anything to the email.
                            Log.Information("Move Email : Read Action set to None. " + email.Sender.Address);
                            break;

                        case EmailAction.Delete: //Delete the email.
                            Log.Information("Move Email : Deleting email from " + email.Sender.Address);
                            email.Delete();
                            break;

                        case EmailAction.Move: //Move the email.
                            Log.Information("Move Email : Moving email from " + email.Sender.Address);
                            MoveEmailToFolder(mailContact.InboxPath, email);
                            break;

                        case EmailAction.Junk:
                            Log.Information("Move Email to Junk: Moving email from " + email.Sender.Address);
                            email.Move(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk));
                            break;
                    }
                }

                mailContact.SaveAndDispose();
            }
            else
            {
                //Get the 'On Behalf' property from the email.
                string onBehalfEmailAddress;

                try
                {
                    Outlook.PropertyAccessor mapiPropertyAccessor;
                    string propertyName = "http://schemas.microsoft.com/mapi/proptag/0x0065001F";
                    mapiPropertyAccessor = email.PropertyAccessor;
                    onBehalfEmailAddress = mapiPropertyAccessor.GetProperty(propertyName).ToString();
                    if (mapiPropertyAccessor is object)
                    {
                        Marshal.ReleaseComObject(mapiPropertyAccessor);
                    }
                }
                catch(Exception ex)
                {
                    Log.Error(ex.Message, ex);
                }

                try
                {
                    if (email is object)
                    {
                        if (email.SenderEmailAddress is object)
                        {
                            Log.Information("Move Email : No Contact for " + email.SenderEmailAddress);
                        }
                        else
                        {
                            Log.Information("Move Email : No Contact (detatched object?)");
                        }
                    }
                    else
                    {
                        Log.Information("Move Email : No email object");
                    }
                    //Op.LogMessage("SenderName         : " + email.SenderName);
                    //Op.LogMessage("SentOnBehalfOfName : " + email.SentOnBehalfOfName);
                    //Op.LogMessage("ReplyRecipientNames: " + email.ReplyRecipientNames);
                    //Op.LogMessage("On Behalf: " + onBehalfEmailAddress);
                    //Op.LogMessage("");
                }
                catch (Exception ex)
                {
                    Log.Error(ex.Message, ex);
                }

                //Log the details.
                
            }

            if (email is object) { Marshal.ReleaseComObject(email); }
            if (contact is object) { Marshal.ReleaseComObject(contact); }

        }

        /// <summary>
        /// Method to move the email from the Inbox to the specified folder.
        /// </summary>
        /// <param name="folderPath">The path to the folder to move the email.</param>
        /// <param name="email">The mailitem to move.</param>
        private static void MoveEmailToFolder(string folderPath, Outlook.MailItem email)
        {
            string[] folders = folderPath.Split('\\');
            Outlook.MAPIFolder folder;
            Outlook.Folders subFolders;

            try
            {
                folder = InTouch.Stores.StoresLookup[folders[0]].RootFolder;
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                Log.Information("Exception managed > Store not found. (" + folders[0] + ")");
                return;
            }

            try
            {
                for (int i = 1; i <= folders.GetUpperBound(0); i++)
                {
                    subFolders = folder.Folders;
                    folder = subFolders[folders[i]] as Outlook.Folder;
                }
            }
            catch (COMException ex)
            {
                if (ex.HResult == -2147221233)
                {
                    Log.Information("Exception Managed > Folder not found. (" + folderPath + ")");
                    return;
                }
                else
                {
                    Log.Error(ex.Message,ex);
                    throw;
                }
            }

            if (folder is object)
            {
                email.Move(folder);
                Marshal.ReleaseComObject(folder);
            }
        }
    }
}