using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InTouch_AutoFile
{
    public partial class FormInTouchNewContact : Form
    {
        private Outlook.MailItem email;
        public Outlook.MailItem Email
        {
            get { return email; }
            set
            {
                email = value;
            }
        }

        private readonly Color colorBackground = SystemColors.Control;
        private readonly Color colorBackgroundError = Color.FromArgb(255, 128, 128);

        public FormInTouchNewContact()
        {
            InitializeComponent();
        }

        private void FormInTouchNewContact_Load(object sender, EventArgs e)
        {
            Op.EmailForCreatedContact = null;
            ComboBoxContactFolder.Items.Add("Contacts");
            foreach (Outlook.Folder nextContactFolder in InTouch.Contacts.Folders)
            {
                switch (nextContactFolder.Name)
                {
                    case "Recipient Cache":
                        break;
                    case "Organizational Contacts":
                        break;
                    case "PeopleCentricConversation Buddies":
                        break;
                    case "GAL Contacts":
                        break;
                    case "{A9E2BC46-B3A0-4243-B315-60D991004455}":
                        break;
                    case "{06967759-274D-40B2-A3EB-D7F9E73727D7}":
                        break;
                    case "Companies":
                        break;
                    default:
                        ComboBoxContactFolder.Items.Add(nextContactFolder.Name);
                        break;
                }
            }
            ComboBoxContactFolder.Text = Properties.Settings.Default.LastContactFolder;
        }

        private void FormInTouchNewContact_Shown(object sender, EventArgs e)
        {
            SetupForm();
        }

        private void SetupForm()
        {
            //Get the 'On Behalf' property from the email.
            Outlook.PropertyAccessor mapiPropertyAccessor;
            string propertyName = "http://schemas.microsoft.com/mapi/proptag/0x0065001F";
            mapiPropertyAccessor = Email.PropertyAccessor;
            string onBehalfEmailAddress = mapiPropertyAccessor.GetProperty(propertyName).ToString();
            if (mapiPropertyAccessor is object) Marshal.ReleaseComObject(mapiPropertyAccessor);

            string website = GetWebSiteAddress(email.SenderEmailAddress);
            GetWebsitesFavicon(website);

            TextBoxFullName.Text = email.SenderName;
            TextBoxEmail1.Text = email.SenderEmailAddress;
            if (email.SenderEmailAddress != onBehalfEmailAddress)
            {
                TextBoxEmail2.Text = onBehalfEmailAddress;
            }
        }

        private static string GetWebSiteAddress(string senderEmailAddress)
        {
            //Get website address.
            string website;
            try
            {
                website = senderEmailAddress;
                website = website.Substring(website.IndexOf("@") + 1);

                if (website.Substring(0, 5).ToLower() == "mail.")
                {
                    website = website.Substring(5);
                }
                else if (website.Substring(0, 7).ToLower() == "mailer.")
                {
                    website = website.Substring(7);
                }
                else if (website.Substring(0, 6).ToLower() == "email.")
                {
                    website = website.Substring(6);
                }

                website = "https://www." + website;
            }
            catch (Exception ex)
            {
                Op.LogError(ex);
                throw;
            }

            Op.LogMessage("Website : " + website);
            return website;
        }

        private void GetWebsitesFavicon(string website)
        {

            if (File.Exists("Icon.png"))
            {
                File.Delete("Icon.png");
            }

            if (website.Length > 0)
            {
                using (var client = new WebClient())
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                    ServicePointManager.Expect100Continue = true;
                    string htmlString = "";

                    try
                    {
                        htmlString = client.DownloadString(website);
                        if (htmlString.Length > 0)
                        {
                            TextBoxWebsite.Text = website;
                        }
                        if (htmlString.IndexOf("<link rel=\"apple - touch - icon\"") >= 0)
                        {
                            htmlString = htmlString.Substring(htmlString.IndexOf("<link rel=\"apple - touch - icon\"") + 32);
                            htmlString = htmlString.Substring(htmlString.IndexOf("href=\"") + 6);
                            htmlString = htmlString.Substring(0, htmlString.IndexOf("\""));
                            Op.LogMessage("HTML String : " + htmlString);
                        }
                        else if (htmlString.IndexOf("<link rel=\"shortcut icon\"") >= 0)
                        {
                            htmlString = htmlString.Substring(htmlString.IndexOf("<link rel=\"shortcut icon\"") + 25);
                            htmlString = htmlString.Substring(htmlString.IndexOf("href=\"") + 6);
                            htmlString = htmlString.Substring(0, htmlString.IndexOf("\""));
                            Op.LogMessage("HTML String : " + htmlString);
                        }
                        else if (htmlString.IndexOf("<link rel=\"icon\"") >= 0)
                        {
                            htmlString = htmlString.Substring(htmlString.IndexOf("<link rel=\"icon\"") + 16);
                            htmlString = htmlString.Substring(htmlString.IndexOf("href=\"") + 6);
                            htmlString = htmlString.Substring(0, htmlString.IndexOf("\""));
                            Op.LogMessage("HTML String : " + htmlString);
                        }
                        else
                        {
                            Op.LogMessage("HTML String : " + htmlString);
                        }

                        if (htmlString.Substring(0, 1) == "/")
                        {
                            htmlString = website + htmlString;
                        }
                    }
                    catch(WebException ex)
                    {
                        switch (ex.HResult)
                        {
                            case -2146233079:
                                Op.LogMessage("Exception Managed > The remote name could not be resolved.");
                                break;

                            default:
                                Op.LogError(ex);
                                throw;
                        }
                    }
                    catch (Exception ex)
                    {
                        Op.LogError(ex);
                        throw;
                    }

                    if (htmlString.Length > 0)
                    {
                        try
                        {
                            client.DownloadFile(htmlString, "Icon.png");
                        }
                        catch (Exception ex)
                        {
                            Op.LogError(ex);
                            throw;
                        }

                        if (File.Exists("Icon.png"))
                        {
                            PictureBoxIcon.Image = Image.FromFile("Icon.png");
                        }
                    }
                }
            }
        }

        private void ButtonCreateContact_Click(object sender, EventArgs e)
        {
            bool noProblems = true;

            //Reset the colors of controls that change.
            TextBoxFullName.BackColor = colorBackground;
            ComboBoxContactFolder.BackColor = colorBackground;
            TextBoxEmail1.BackColor = colorBackground;
            TextBoxEmail2.BackColor = colorBackground;
            TextBoxEmail3.BackColor = colorBackground;

            //Check for required fields.
            if (TextBoxFullName.Text.Length == 0)
            {
                TextBoxFullName.BackColor = colorBackgroundError;
                noProblems = false;
            }

            if (ComboBoxContactFolder.Text.Length == 0)
            {
                ComboBoxContactFolder.BackColor = colorBackgroundError;
                noProblems = false;
            }

            //Check if there is at least one email.
            if ((TextBoxEmail1.Text.Length == 0) && (TextBoxEmail1.Text.Length == 0) && (TextBoxEmail1.Text.Length == 0))
            {
                TextBoxEmail1.BackColor = colorBackgroundError;
                noProblems = false;
            }

            //Check the email addresses are already in use.
            if (TextBoxEmail1.Text.Length > 0)
            {
                if (Contacts.DoesLookupContain(TextBoxEmail1.Text))
                {
                    TextBoxEmail1.BackColor = colorBackgroundError;
                    noProblems = false;
                }
            }
            if (TextBoxEmail2.Text.Length > 0)
            {
                if (Contacts.DoesLookupContain(TextBoxEmail2.Text))
                {
                    TextBoxEmail2.BackColor = colorBackgroundError;
                    noProblems = false;
                }
            }
            if (TextBoxEmail3.Text.Length > 0)
            {
                if (Contacts.DoesLookupContain(TextBoxEmail3.Text))
                {
                    TextBoxEmail3.BackColor = colorBackgroundError;
                    noProblems = false;
                }
            }

            //If all requirement are met then create the contact and add the details.
            if (noProblems)
            {
                Outlook.ContactItem newContact = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
                newContact.FullName = TextBoxFullName.Text;
                if (TextBoxEmail1.Text.Length > 0)
                {
                    newContact.Email1Address = TextBoxEmail1.Text;

                }
                if (TextBoxEmail2.Text.Length > 0)
                {
                    newContact.Email2Address = TextBoxEmail2.Text;

                }
                if (TextBoxEmail3.Text.Length > 0)
                {
                    newContact.Email3Address = TextBoxEmail3.Text;
                }

                if(PictureBoxIcon.Image is object)
                {
                    newContact.AddPicture("Icon.png");
                }
                newContact.Save();
                Application.DoEvents();
                Thread.Sleep(100);
                Op.EmailForCreatedContact = newContact.Email1Address;
                if (ComboBoxContactFolder.Text != "Contacts")
                {
                    newContact.Move(InTouch.Contacts.Folders[ComboBoxContactFolder.Text]); ;
                }
                Op.NextFormRegion = ContactFormRegion.InTouchSettings;
                if (newContact is object) { Marshal.ReleaseComObject(newContact); }
            }

            if (noProblems) { Close(); }
        }

        private void FormInTouchNewContact_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.LastContactFolder = ComboBoxContactFolder.Text;
            Properties.Settings.Default.Save();

            if (email is object) { Marshal.ReleaseComObject(email); }
        }

        
    }
}
