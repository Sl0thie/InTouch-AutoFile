namespace InTouch_AutoFile
{
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Runtime.InteropServices;
    using Serilog;

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

        /// <summary>
        /// If True, then when contacts display the InTouch setting form region should show first.
        /// </summary>
        public static bool ShowInTouchSettings
        {
            get { return showInTouchSettings; }
            set { showInTouchSettings = value; }
        }
        private static bool showInTouchSettings = false;
        
        /// <summary>
        /// If True then office is using a dark theme.
        /// </summary>
        public static bool DarkTheme
        {
            get { return darkTheme; }
            set { darkTheme = value; }
        }
        private static bool darkTheme;

        /// <summary>
        /// Checks if the supplied path is a valid path within the inbox branch.
        /// </summary>
        /// <param name="folderPath">The path to test.</param>
        /// <returns>True if the path is valid/False if it is not.</returns>
        public static bool CheckFolderPath(string folderPath)
        {
            bool returnValue = true;
            string[] folders = folderPath.Split('\\');
            Outlook.MAPIFolder folder = null;
            Outlook.Folders subFolders;

            try
            {
                folder = InTouch.Stores.StoresLookup[folders[0]].RootFolder;
            }
            catch (System.Collections.Generic.KeyNotFoundException)
            {
                Log.Information("Exception managed > Store not found. (" + folders[0] + ")");
                returnValue = false;
            }

            if (returnValue)
            {
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
                        Log.Error(ex.Message, ex);
                        returnValue = false;
                        Log.Information("Exception Managed > Folder not found. (" + folderPath + ")");
                    }
                    else
                    {
                        Log.Error(ex.Message, ex);
                        throw;
                    }
                }
            }

            if (folder is object) { Marshal.ReleaseComObject(folder); }

            return returnValue;
        }

        public static void CreatePath(string folderPath)
        {

            string[] folders = folderPath.Split('\\');
            Outlook.MAPIFolder folder;
            Outlook.Folders subFolders;

            try
            {
                folder = InTouch.Stores.StoresLookup[folders[0]].RootFolder;
            }
            catch (System.Collections.Generic.KeyNotFoundException ex)
            {
                Log.Error(ex.Message, ex);
                throw;
                //Op.LogMessage("Exception managed > Store not found. (" + folders[0] + ")");
                //return;
            }

            for (int i = 1; i <= folders.GetUpperBound(0); i++)
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
                        Log.Information("Exception Managed > Creating Folder.");
                        folder.Folders.Add(folders[i]);
                        folder = subFolders[folders[i]] as Outlook.Folder;
                    }

                }

            }
            //}
            //catch (COMException ex)
            //{
            //    if (ex.HResult == -2147221233)
            //    {
            //        //returnValue = false;
            //        //Op.LogMessage("Exception Managed > Folder not found. (" + folderPath + ")");

            //        Op.LogMessage("Exception Managed > Creating Folder.");

            //        folder.Folders.Add(folders[i]);

            //        folder = subFolders[folders[i]] as Outlook.Folder;

            //    }
            //    else
            //    {
            //        throw;
            //    }
            //}

            if (folder is object) { Marshal.ReleaseComObject(folder); }

            //Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(rootFolderType) as Outlook.Folder;
            //Outlook.Folders subFolders;
            //string[] folders = folderPath.Split('\\');

            //for (int i = 0; i <= folders.GetUpperBound(0); i++)
            //{
            //    subFolders = folder.Folders;

            //    try
            //    {
            //        folder = subFolders[folders[i]] as Outlook.Folder;
            //    }
            //    catch (COMException ex)
            //    {
            //        if (ex.HResult == -2147221233)
            //        {
            //            Op.LogMessage("Exception Managed > Creating Folder.");

            //            folder.Folders.Add(folders[i]);

            //            folder = subFolders[folders[i]] as Outlook.Folder;
            //        }
            //        else
            //        {
            //            throw;
            //        }
            //    }

            //}

            //if (folder is object) { Marshal.ReleaseComObject(folder); }
        }

    }
}

