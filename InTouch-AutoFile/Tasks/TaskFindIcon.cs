namespace InTouch_AutoFile.Tasks
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Threading;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using Serilog;
    using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

    internal class TaskFindIcon
    {
        private readonly Action callBack;

        public TaskFindIcon(Action callBack)
        {
            this.callBack = callBack;
        }

        Outlook.Folder contactsFolderOthers = null;

        public void RunTask()
        {
            // If task is enabled in the settings then start task.
            if (Properties.Settings.Default.TaskInbox)
            {
                Log.Information("Starting FindIcon Task.");
                Thread backgroundThread = new Thread(new ThreadStart(BackgroundProcess))
                {
                    Name = "AF.FindIcon",
                    IsBackground = true,
                    Priority = ThreadPriority.Normal
                };
                backgroundThread.SetApartmentState(ApartmentState.STA);
                backgroundThread.Start();
            }
        }

        private void BackgroundProcess()
        {
            DateTime lastDate = Properties.Settings.Default.LastAliasCheck;

            ProcessOtherContacts();

            callBack?.Invoke();
        }

        private void ProcessOtherContacts()
        {
            try
            {
                contactsFolderOthers = (Outlook.Folder)Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders[InTouch.OtherFolderName];
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }

            try
            {
                foreach (object nextObject in contactsFolderOthers.Items)
                {
                    if (nextObject is Outlook.ContactItem contact)
                    {
                        ProcessContact(contact);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }

        }

        private void ProcessContact(Outlook.ContactItem contact)
        {
            string website = "";

            if(contact.Email1Address is object)
            {
                website = GetWebSiteAddress(contact.Email1Address);
            }
            else if (contact.Email2Address is object)
            {
                website = GetWebSiteAddress(contact.Email2Address);
            }
            else if (contact.Email3Address is object)
            {
                website = GetWebSiteAddress(contact.Email3Address);
            }

            if(website != "")
            {
                Log.Information($"{contact.FullName} {website}");
            }
        }

        private static string GetWebSiteAddress(string senderEmailAddress)
        {
            // Get website address.
            string website;
            try
            {
                website = senderEmailAddress.ToLower();
                website = website.Substring(website.IndexOf("@") + 1);

                if (website.Substring(0, 5) == "mail.")
                {
                    website = website.Substring(5);
                }
                else if (website.Substring(0, 7) == "mailer.")
                {
                    website = website.Substring(7);
                }
                else if (website.Substring(0, 6) == "email.")
                {
                    website = website.Substring(6);
                }
                else if (website.Substring(0, 7) == "e-mail.")
                {
                    website = website.Substring(7);
                }
                else if (website.Substring(0, 2) == "e.")
                {
                    website = website.Substring(2);
                }

                switch (website)
                {
                    case "gmail.com":
                    case "outlook.com":
                    case "outlook.com.au":
                    case "hotmail.com":
                        return "";
                        break;
                }

                website = "https://www." + website;
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
                return "";
            }

            //Log.Information("Website : " + website);
            return website;
        }
    }
}
