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

    internal class TaskMonitorAliases
    {
        private readonly Action callBack;

        public TaskMonitorAliases(Action callBack)
        {
            this.callBack = callBack;
        }

        public void RunTask()
        {
            // If task is enabled in the settings then start task.
            if (Properties.Settings.Default.TaskInbox)
            {
                Log.Information("Starting MonitorAliases Task.");
                Thread backgroundThread = new Thread(new ThreadStart(BackgroundProcess))
                {
                    Name = "AF.MonitorAliases",
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

            if(DateTime.Now > lastDate.AddDays(1))
            {
                Properties.Settings.Default.CurrentAliasGUID = Guid.NewGuid().ToString();
                Properties.Settings.Default.LastAliasCheck = DateTime.Now;
                Properties.Settings.Default.Save();

                ProcessAliases();
            }

            callBack?.Invoke();
        }

        private void ProcessAliases()
        {
            Outlook.MAPIFolder aliasesFolder = null;

            try
            {
                aliasesFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders[InTouch.AliasFolderName];
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
                Log.Information($"Can't find {InTouch.AliasFolderName} folder.");
                return;
            }

            try
            {
                foreach (object nextObject in aliasesFolder.Items)
                {
                    if (nextObject is Outlook.ContactItem contact)
                    {
                        Log.Information($"Alias {((Outlook.ContactItem)nextObject).FullName}");

                        if(((Outlook.ContactItem)nextObject).Email1Address is object)
                        {
                            if (((Outlook.ContactItem)nextObject).Email1Address.Trim() != "")
                            {
                                SendEmail(((Outlook.ContactItem)nextObject).Email1Address);
                            }
                        }

                        if (((Outlook.ContactItem)nextObject).Email2Address is object)
                        {
                            if (((Outlook.ContactItem)nextObject).Email2Address.Trim() != "")
                            {
                                SendEmail(((Outlook.ContactItem)nextObject).Email2Address);
                            }
                        }

                        if (((Outlook.ContactItem)nextObject).Email3Address is object)
                        {
                            if (((Outlook.ContactItem)nextObject).Email3Address.Trim() != "")
                            {
                                SendEmail(((Outlook.ContactItem)nextObject).Email3Address);
                            }
                        }
                    }

                    if (nextObject is object)
                    {
                        Marshal.ReleaseComObject(nextObject);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }
        }

        private void SendEmail(string address)
        {
            Log.Information($"Sending Email to {address}");

            Outlook.MailItem eMail = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = $"ALIAS PATH CHECK [{Properties.Settings.Default.CurrentAliasGUID}]";
            eMail.To = address;
            eMail.Body = $"ALIAS PATH CHECK [{Properties.Settings.Default.CurrentAliasGUID}]";
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            ((Outlook._MailItem)eMail).Send();
        }
    }
}
