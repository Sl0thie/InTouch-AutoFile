using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace InTouch_AutoFile
{
    public static class InTouch
    {
        
        public static Contacts Contacts
        {
            get { return contacts; }
        }
        private static readonly Contacts contacts = new Contacts();
        
        public static TaskManager TaskManager
        {
            get { return taskManager; }
        }
        private static readonly TaskManager taskManager = new TaskManager();


        public static Stores Stores
        {
            get { return stores; }
            set { stores = value; }
        }
        private static Stores stores = new Stores();






        public static void CreatePath(string folderPath, Outlook.OlDefaultFolders rootFolderType)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(rootFolderType) as Outlook.Folder;
            Outlook.Folders subFolders;
            string[] folders = folderPath.Split('\\');

            for (int i = 0; i <= folders.GetUpperBound(0); i++)
            {
                subFolders = folder.Folders;

                try
                {
                    folder = subFolders[folders[i]] as Outlook.Folder;
                }
                catch (COMException ex)
                {
                    if (ex.HResult == -2147221233)
                    {
                        Op.LogMessage("Exception Managed > Creating Folder.");

                        folder.Folders.Add(folders[i]);

                        folder = subFolders[folders[i]] as Outlook.Folder;
                    }
                    else
                    {
                        throw;
                    }
                }

            }
        }




    }
}
