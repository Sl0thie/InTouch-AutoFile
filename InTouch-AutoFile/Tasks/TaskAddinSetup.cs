namespace InTouch_AutoFile
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using Serilog;

    internal class TaskAddinSetup
    {
        private readonly Action callBack;

        public TaskAddinSetup(Action callBack)
        {
            this.callBack = callBack;
        }

        public void RunTask()
        {
            // If task is enabled in the settings then start task.
            if (Properties.Settings.Default.TaskInbox)
            {
                Log.Information("Starting AddinSetup Task.");
                Thread backgroundThread = new Thread(new ThreadStart(BackgroundProcess))
                {
                    Name = "AF.AddinSetup",
                    IsBackground = true,
                    Priority = ThreadPriority.Normal
                };
                backgroundThread.SetApartmentState(ApartmentState.STA);
                backgroundThread.Start();
            }
        }

        private void BackgroundProcess()
        {
            SetupAddin();
            callBack?.Invoke();
        }

        private void SetupAddin()
        {
            // Check if contacts stores exist. If not create them.
            Outlook.MAPIFolder contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts); ;
            Outlook.MAPIFolder contactsFolderJunk = null;
            Outlook.MAPIFolder contactsFolderOthers = null;

            try
            {
                contactsFolderJunk = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders[InTouch.JunkFolderName];
            }
            catch(Exception ex)
            {
                Log.Error(ex.Message, ex);
            }

            try
            {
                if (contactsFolderJunk is null)
                {
                    //Log.Information($"Creating {InTouch.JunkFolderName} folder.");
                    contactsFolder.Folders.Add(InTouch.JunkFolderName);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }

            try
            {
                contactsFolderOthers = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders[InTouch.OtherFolderName];
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }

            try
            {
                if (contactsFolderJunk is null)
                {
                    //Log.Information($"Creating {InTouch.OtherFolderName} folder.");
                    contactsFolder.Folders.Add(InTouch.OtherFolderName);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }

            //Outlook.MAPIFolder contactsFolderJunk = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders["Junk Contacts"];

            //Outlook.MAPIFolder contactsFolderOthers = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders["Others"];

            //if(contactsFolderOthers is null)
            //{
            //    contactsFolder.Folders.Add("Others");
            //}

            //if (contactsFolderJunk is null)
            //{
            //    contactsFolder.Folders.Add("Junk Contacts");
            //}
        }
    }
}
