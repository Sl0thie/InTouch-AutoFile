using System;
using System.Threading;
using System.Diagnostics;
using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Collections;

namespace InTouch_AutoFile
{
    

    public static class Log
    {
        /// <summary>
        /// Method to centralise message logging. This is to provide a standard message format.
        /// </summary>
        /// <param name="message">The Message to be logged.</param>
        public static void Message(string message)
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
        public static void Error(Exception ex)
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

    }
}
