using System;
using System.Threading;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace InTouch_AutoFile
{
    #region Enumerations

    public enum ContactFormRegion
    {
        None, InTouchSettings
    }
    public enum EmailAction
    {
        None, Delete, Move
    }

    #endregion

    public static class Op
    {
        private static ContactFormRegion nextFormRegion = ContactFormRegion.None;
        public static ContactFormRegion NextFormRegion
        {
            get { return nextFormRegion; }
            set { nextFormRegion = value; }
        }

        private static string contactCreatedEmail;
        //TODO Rename this to suit what it has become.
        public static string ContactCreatedEmail
        {
            get { return contactCreatedEmail; }
            set { contactCreatedEmail = value; }
        }

        /// <summary>
        /// Checks if the supplied path is a valid path within the inbox branch.
        /// </summary>
        /// <param name="folderPath">The path to test.</param>
        /// <returns>True if the path is valid/False if it is not.</returns>
        public static bool CheckFolderPath(string folderPath, Outlook.OlDefaultFolders rootFolderType)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(rootFolderType) as Outlook.Folder;
            string backslash = @"\";

            if (folderPath is null)
            {
                throw new Exception("The parameter 'folderPath' cannot be null.");
            }

            String[] folders = folderPath.Split(backslash.ToCharArray());

            if (folder is object)
            {
                for (int i = 0; i < folders.GetUpperBound(0); i++)
                {
                    Outlook.Folders subFolders = folder.Folders;
                    if (subFolders is object)
                    {
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null)
                        {
                            if (folder is object) { Marshal.ReleaseComObject(folder); }
                            return false;
                        }
                    }
                    else
                    {
                        if (folder is object) { Marshal.ReleaseComObject(folder); }
                        return false;
                    }
                }
            }
            else
            {
                throw new Exception("Root folder is null");
            }

            if (folder is object) { Marshal.ReleaseComObject(folder); }
            return true;
        }

        #region Logging Methods

        /// <summary>
        /// Method to centralise message logging. This is to provide a standard message format.
        /// </summary>
        /// <param name="message">The Message to be logged.</param>
        public static void LogMessage(string message)
        {
            if (message is object)
            {
                if (Thread.CurrentThread is object)
                {
                    if (Thread.CurrentThread.Name is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + Thread.CurrentThread.Name.PadRight(30) + " : INFO      : " + message);
                    }
                    else
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + "Unknown Thread    " + " : INFO      : " + message);
                    }
                }
                else
                {
                    Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + "Unknown Thread    " + " : INFO      : " + message);
                }
            }
        }

        /// <summary>
        /// Method to centralise error logging.
        /// </summary>
        /// <param name="ex">The Exception to be logged.</param>
        public static void LogError(Exception ex)
        {
            if (ex is object)
            {
                if (Thread.CurrentThread.Name is object)
                {
                    if (ex.Message is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + Thread.CurrentThread.Name.PadRight(30) + " : EXCEPTION :" + ex.Message);
                    }
                    else
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + Thread.CurrentThread.Name.PadRight(30) + " Message : No Message");
                    }

                    if (ex.Source is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + Thread.CurrentThread.Name.PadRight(30) + " Source : " + ex.Source);
                    }

                    if (ex.StackTrace is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + Thread.CurrentThread.Name.PadRight(30) + " StackTrace : " + ex.StackTrace);
                    }

                }
                else
                {
                    if (ex.Message is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + "Unknown Thread    " + " : EXCEPTION :" + ex.Message);
                    }
                    else
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + "Unknown Thread    " + " Message : No Message");
                    }

                    if (ex.Source is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + "Unknown Thread    " + " Source : " + ex.Source);
                    }

                    if (ex.StackTrace is object)
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + " : " + "Unknown Thread    " + " StackTrace : " + ex.StackTrace);
                    }

                }
            }
        }

        #endregion
    }
}
