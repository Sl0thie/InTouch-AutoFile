using System;
using System.Threading;
using System.Diagnostics;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Collections;

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

        private static string emailForCreatedContact;
        public static string EmailForCreatedContact
        {
            get { return emailForCreatedContact; }
            set { emailForCreatedContact = value; }
        }
        /// <summary>
        /// Checks if the supplied path is a valid path within the inbox branch.
        /// </summary>
        /// <param name="folderPath">The path to test.</param>
        /// <returns>True if the path is valid/False if it is not.</returns>
        public static bool CheckFolderPath(string folderPath, Outlook.OlDefaultFolders rootFolderType)
        {
            bool returnValue = true;
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(rootFolderType) as Outlook.Folder;
            Outlook.Folders subFolders;
            string[] folders = folderPath.Split('\\');

            try
            {
                for (int i = 0; i <= folders.GetUpperBound(0); i++)
                {
                    subFolders = folder.Folders;
                    folder = subFolders[folders[i]] as Outlook.Folder;
                }
            }
            catch(COMException ex)
            {
                if(ex.HResult == -2147221233)
                {
                    returnValue = false;
                    Op.LogMessage("Exception Managed.");
                }
                else
                {
                    throw;
                }
            }
            if (folder is object) { Marshal.ReleaseComObject(folder); }
            return returnValue;
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
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + ": " + Thread.CurrentThread.Name.PadRight(15) + ": INFO          : " + message);
                    }
                    else
                    {
                        Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + ": " + "Unknown Thread " + ": INFO          : " + message);
                    }
                }
                else
                {
                    Debug.WriteLine(DateTime.Now.ToString("h:mm:ss:fff") + ": " + "Unknown Thread " + ": INFO          : " + message);
                }
            }
        }

        /// <summary>
        /// Method to centralise error logging.
        /// </summary>
        /// <param name="ex">The Exception to be logged.</param>
        public static void LogError(Exception ex)
        {
            if(ex is object)
            {
                string timeAndThreadName = DateTime.Now.ToString("h:mm:ss:fff");
                
                if (Thread.CurrentThread.Name is object)
                {
                    timeAndThreadName += ": " + Thread.CurrentThread.Name.PadRight(15) + ": ";
                }
                else
                {
                    timeAndThreadName += ": Unknown Thread : ";
                }

                Debug.WriteLine(timeAndThreadName + "EXCEPTION     : ");


                if (ex.Message is object)
                {
                    Debug.WriteLine(timeAndThreadName + "Message       : " + ex.Message);
                }

                Debug.WriteLine(timeAndThreadName + "HResult       : " + ex.HResult);

                if (ex.Source is object)
                {
                    Debug.WriteLine(timeAndThreadName + "Source        : " + ex.Source);
                }

                if(ex.TargetSite is object)
                {
                    Debug.WriteLine(timeAndThreadName + "TargetSite    : " + ex.TargetSite.Name);
                    
                    //ParameterInfo[] parameters = ex.TargetSite.GetParameters();
                    //if(parameters.Length > 0)
                    //{
                    //    Debug.WriteLine(timeAndThreadName + "TargetSite.Parameters (" + parameters.Length + ")");
                    //    foreach (ParameterInfo parameter in parameters)
                    //    {
                    //        Debug.WriteLine(timeAndThreadName + "    " + parameter.Name + " " + parameter.ParameterType.FullName);
                    //    }
                    //}
                }

                if((ex.Data is object) && (ex.Data.Count > 0))
                {
                    Debug.WriteLine(timeAndThreadName + "Data : ");
                    foreach(DictionaryEntry nextPair in ex.Data)
                    {
                        Debug.WriteLine(timeAndThreadName + "Key : " + nextPair.Key.ToString() + " value : " + nextPair.Value.ToString());
                    }
                }

                if (ex.StackTrace is object)
                {
                    string[] lines = ex.StackTrace.Split('\n');
                    if(lines.Length > 0)
                    {
                        Debug.WriteLine(timeAndThreadName + "StackTrace    :");
                        foreach (string nextLine in lines)
                        {
                            string line = nextLine.Replace('\n', ' ');
                            line = line.Replace('\r', ' ');
                            line = line.Trim();
                            Debug.WriteLine(timeAndThreadName + "              : " + line);
                        }
                    }
                }

                if (ex.InnerException is object)
                {
                    Debug.WriteLine(timeAndThreadName + "InnerException: ");

                    if (ex.InnerException.Message is object)
                    {
                        Debug.WriteLine(timeAndThreadName + "               Message : " + ex.InnerException.Message);
                    }

                    Debug.WriteLine(timeAndThreadName + "               HResult : " + ex.InnerException.HResult);

                    if (ex.InnerException.Source is object)
                    {
                        Debug.WriteLine(timeAndThreadName + "               Source : " + ex.InnerException.Source);
                    }

                    if (ex.InnerException.TargetSite is object)
                    {
                        Debug.WriteLine(timeAndThreadName + "               TargetSite : " + ex.InnerException.TargetSite.Name);
                    }

                    if (ex.InnerException.Data is object)
                    {
                        Debug.WriteLine(timeAndThreadName + "               Data : ");
                        foreach (DictionaryEntry nextPair in ex.InnerException.Data)
                        {
                            Debug.WriteLine(timeAndThreadName + "               Key : " + nextPair.Key.ToString() + " value : " + nextPair.Value.ToString());
                        }
                    }

                    if (ex.InnerException.StackTrace is object)
                    {
                        Debug.WriteLine(timeAndThreadName + "               StackTrace : " + ex.InnerException.StackTrace);
                    }
                }

                Debug.WriteLine("");
            }
        }

        #endregion
    }
}
