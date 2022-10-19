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
    using System.Net;
    using System.IO;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.Windows.Media.Imaging;

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

            if (!contact.HasPicture)
            {
                string website = "";

                if (contact.Email1Address is object)
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

                if (website != "")
                {
                    Log.Information($"{contact.FullName} {website}");

                    GetWebsitesFavicon(website);

                    if (File.Exists("Icon.png"))
                    {
                        contact.AddPicture("Icon.png");
                        contact.Save();
                    }
                }
            }

            if (contact is object)
            {
                Marshal.ReleaseComObject(contact);
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

        private void GetWebsitesFavicon(string website)
        {
            // Delete temp icon file.
            if (File.Exists("Icon.png"))
            {
                File.Delete("Icon.png");
            }

            if (website.Length > 0)
            {
                using (WebClient client = new WebClient())
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                    ServicePointManager.Expect100Continue = true;
                    string htmlString = "";

                    try
                    {
                        htmlString = client.DownloadString(website);

                        if (htmlString.IndexOf("<link rel=\"apple - touch - icon\"") >= 0)
                        {
                            htmlString = htmlString.Substring(htmlString.IndexOf("<link rel=\"apple - touch - icon\"") + 32);
                            htmlString = htmlString.Substring(htmlString.IndexOf("href=\"") + 6);
                            htmlString = htmlString.Substring(0, htmlString.IndexOf("\""));
                            //Log.Information("HTML String : " + htmlString);
                        }
                        else if (htmlString.IndexOf("<link rel=\"shortcut icon\"") >= 0)
                        {
                            htmlString = htmlString.Substring(htmlString.IndexOf("<link rel=\"shortcut icon\"") + 25);
                            htmlString = htmlString.Substring(htmlString.IndexOf("href=\"") + 6);
                            htmlString = htmlString.Substring(0, htmlString.IndexOf("\""));
                            //Log.Information("HTML String : " + htmlString);
                        }
                        else if (htmlString.IndexOf("<link rel=\"icon\"") >= 0)
                        {
                            htmlString = htmlString.Substring(htmlString.IndexOf("<link rel=\"icon\"") + 16);
                            htmlString = htmlString.Substring(htmlString.IndexOf("href=\"") + 6);
                            htmlString = htmlString.Substring(0, htmlString.IndexOf("\""));
                            //Log.Information("HTML String : " + htmlString);
                        }
                        else
                        {
                            //Log.Information("HTML String : " + htmlString);
                        }

                        if (htmlString.Substring(0, 1) == "/")
                        {
                            htmlString = website + htmlString;
                        }

                        Log.Information("HTML String : " + htmlString);
                    }
                    catch (WebException ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233079:
                                Log.Information("Exception Managed > The remote name could not be resolved.");
                                break;

                            default:
                                Log.Error(ex.Message, ex);
                                //throw;
                                return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex.Message, ex);
                        //throw;
                        return;
                    }

                    if (htmlString.Length > 0)
                    {
                        try
                        {
                            //client.DownloadFile(htmlString, "Icon.png");

                            client.DownloadFile(htmlString, "Icon.ico");

                           
                            Icon i = new Icon("Icon.ico");
                            Bitmap bm = PngFromIcon(i);

                            bm.Save("Icon.png", ImageFormat.Png);
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            //throw;
                        }
                    }
                }
            }
        }

        public static Bitmap PngFromIcon(Icon icon)
        {
            Bitmap png = null;
            using (var iconStream = new System.IO.MemoryStream())
            {
                icon.Save(iconStream);
                var decoder = new IconBitmapDecoder(iconStream,
                    BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.None);

                using (var pngSteam = new System.IO.MemoryStream())
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(decoder.Frames[0]);
                    encoder.Save(pngSteam);
                    png = (Bitmap)Bitmap.FromStream(pngSteam);
                }
            }
            return png;
        }

    }
}
