using System;
using System.Resources;
using Outlook = Microsoft.Office.Interop.Outlook;
using Serilog;

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
            // Setup logging for the application.
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Debug()
                .WriteTo.File("InTouch-Autofile - .txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            Log.Information("Add-in Startup");

            // Add event handlers for when mail is incoming or outgoing.
            Application.NewMailEx += Application_NewMailEx;
            Application.ItemSend += Application_ItemSend;

            // Event handler for the Outlook Add-in's Options page.
            Application.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);

            // Queue up tasks to start up after launch.
            InTouch.TaskManager.EnqueueInboxTask();
            InTouch.TaskManager.EnqueueSentItemsTask();

            GetRegistrySettings();
        }

        private static void GetRegistrySettings()
        {
            object uIThemeObj = Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common","UI Theme",null);
            uint uITheme = Convert.ToUInt32(uIThemeObj);
            
            switch (uITheme)
            {
                case 0:// Colorful
                case 3:// Dark gray
                    InTouch.DarkTheme = false;
                    break;
                case 4:// Black
                    InTouch.DarkTheme = true;
                    break;
                case 5:// White
                    InTouch.DarkTheme = false;
                    break;
                case 6:// System Settings
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
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }

        #endregion
    }
}