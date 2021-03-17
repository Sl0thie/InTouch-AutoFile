using System;
using System.Resources;
using Outlook = Microsoft.Office.Interop.Outlook;

[assembly: CLSCompliant(false)]
[assembly: NeutralResourcesLanguage("en-US")]
namespace InTouch_AutoFile
{
    #region Enumerations

    public enum EmailAction
    {
        None, Delete, Move, Junk
    }

    #endregion

    public partial class ThisAddIn
    {

        //TODO Make ribbon icon change color to suit theme.

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //Add event handlers for when mail is incoming or outgoing.
            Application.NewMailEx += Application_NewMailEx;
            Application.ItemSend += Application_ItemSend;

            //Event handler for the Outlook Add-in's Options page.
            Application.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);

            //Queue up tasks to start up after launch.
            InTouch.TaskManager.EnqueueInboxTask();
            InTouch.TaskManager.EnqueueSentItemsTask();

            GetRegistrySettings();           
        }

        private void GetRegistrySettings()
        {
            object uIThemeObj = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common","UI Theme",null);
            UInt32 uITheme = System.Convert.ToUInt32(uIThemeObj);
            //Op.LogMessage("UI Theme : " + uITheme.ToString());
            switch (uITheme)
            {
                case 0://Colorful
                case 3://Darkgrey
                    InTouch.DarkTheme = false;
                    break;
                case 4://Black
                    InTouch.DarkTheme = true;
                    break;
                case 5://White
                    InTouch.DarkTheme = false;
                    break;
                case 6://System Settings
                    //TODO Check the system setting for the color. (High Contrast)
                    InTouch.DarkTheme = false;
                    break;
                default:
                    InTouch.DarkTheme = false;
                    break;
            }
        }

        private void Application_OptionsPagesAdd(Outlook.PropertyPages pages)
        {
            pages.Add(new UCPropertPage(), "");
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            InTouch.TaskManager.EnqueueSentItemsTask();
        }

        private void Application_NewMailEx(string entryIDCollection)
        {
            InTouch.TaskManager.EnqueueInboxTask();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }


        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }

        #endregion
    }
}