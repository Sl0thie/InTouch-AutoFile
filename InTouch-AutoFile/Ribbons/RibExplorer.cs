namespace InTouch_AutoFile
{
    using Microsoft.Office.Tools.Ribbon;
    using System;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Drawing;
    using System.Windows.Forms;
    using System.Threading;
    using System.Threading.Tasks;
    using Serilog;

    //TODO Redo the task to include cancellation so they don't run into each other.
    /// <summary>
    /// A Ribbon extension for the Outlook Explorer Menu to provide buttons for InTouch.
    /// </summary>
    public partial class RibExplorer
    {
        private Outlook.Explorer explorer; //The current explorer object.

        private string lastEntryID = "";

        /// <summary>
        /// Constructor for the explorer ribbon.
        /// This ribbon show when emails are selected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RibExplorer_Load(object sender, RibbonUIEventArgs e)
        {
            explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            explorer.SelectionChange += Explorer_SelectionChange;

            //Fire for first event that is missed during startup.
            Task.Factory.StartNew(() => { CheckEmailSender(); });
            //Parallel.Invoke(() => { CheckEmailSender(); });
        }

        /// <summary>
        /// Event handler for when the item is changed in outlook. 
        /// If it is a email then manage the buttons on the ribbon.
        /// </summary>
        private void Explorer_SelectionChange()
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                //check the first object selected.
                dynamic selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];

                if (selectedObject.EntryID != lastEntryID)
                {
                    lastEntryID = selectedObject.EntryID;
                    ClearButtons();

                    //If it's a mail object then manage based on the contacts details.
                    if (selectedObject is Outlook.MailItem)
                    {
                        Task.Factory.StartNew(() => { CheckEmailSender(); });
                        //Parallel.Invoke(() => { CheckEmailSender(); });
                    }
                    else if (selectedObject is Outlook.TaskItem) { lastEntryID = ""; }
                    else if (selectedObject is Outlook.ContactItem) { lastEntryID = ""; }
                    else if (selectedObject is Outlook.AppointmentItem) { lastEntryID = ""; }
                }

                if (selectedObject is object) Marshal.ReleaseComObject(selectedObject);
            }
            else
            {
                ClearButtons();
            }
        }

        private static void ClearButtons()
        {
            //Clear all the buttons.
            Globals.Ribbons.RibExplorer.buttonContact.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAddContactPersonal.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAddContactOther.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAddContactJunk.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAttention.Visible = false;

            Application.DoEvents();
        }

        private void CheckEmailSender()
        {
            Outlook.Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            Outlook.MailItem email = null;
            if (selection.Count > 0)
            {
                email = selection[1] as Outlook.MailItem;
            }

            if (email is object)
            {
                if (email.Sender is object)
                {
                    //Try to find contact from email address.
                    InTouchContact emailContact = null;
                    try
                    {
                        Outlook.ContactItem contact = InTouch.Contacts.FindContactFromEmailAddress(email.Sender.Address);
                        if (contact is object)
                        {
                            emailContact = new InTouchContact(contact);
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex.Message, ex);
                        throw;
                    }

                    if (emailContact is object)
                    {
                        //Make the Contact Button visible and add the image and name to the button.
                        Globals.Ribbons.RibExplorer.buttonContact.Visible = true;

                        if (emailContact.FullName is object)
                        {
                            Globals.Ribbons.RibExplorer.buttonContact.Label = emailContact.FullName;
                        }
                        else
                        {
                            Globals.Ribbons.RibExplorer.buttonContact.Label = "";
                        }

                        if (emailContact.HasPicture)
                        {
                            Globals.Ribbons.RibExplorer.buttonContact.Image = Image.FromFile(emailContact.GetContactPicturePath());
                        }
                        else
                        {
                            Globals.Ribbons.RibExplorer.buttonContact.Image = Properties.Resources.contact;
                        }

                        //Check if the contact details are valid.
                        if (!emailContact.CheckDetails())
                        {
                            Globals.Ribbons.RibExplorer.buttonAttention.Visible = true;
                        }
                        emailContact.SaveAndDispose();
                    }
                    else
                    {
                        //As the contact was not found, make the add contact button visible.
                        Globals.Ribbons.RibExplorer.buttonAddContactPersonal.Visible = true;
                        Globals.Ribbons.RibExplorer.buttonAddContactOther.Visible = true;
                        Globals.Ribbons.RibExplorer.buttonAddContactJunk.Visible = true;
                    }
                }
                else
                {
                    Globals.Ribbons.RibExplorer.buttonAddContactPersonal.Visible = true;
                    Globals.Ribbons.RibExplorer.buttonAddContactOther.Visible = true;
                    Globals.Ribbons.RibExplorer.buttonAddContactJunk.Visible = true;
                }

                //if (email.EntryID != lastEntryID)
                //{
                //    //Track last EntryID as Explorer_SelectionChange event fires twice for each selection change.
                //    lastEntryID = email.EntryID;

                //    ClearButtons();

                    
                //}              
            }

            if (email is object) { Marshal.ReleaseComObject(email); }
            if (selection is object) { Marshal.ReleaseComObject(selection); }
        }

        private void ButtonContact_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    Outlook.ContactItem emailContact;
                    try
                    {
                        emailContact = InTouch.Contacts.FindContactFromEmailAddress(email.Sender.Address);
                    }
                    catch (Exception ex) 
                    { 
                        Log.Error(ex.Message,ex); 
                        throw; 
                    }

                    if (emailContact is object)
                    {
                        emailContact.Display(false);
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                    if (emailContact is object) { Marshal.ReleaseComObject(emailContact); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
            }
        }

        private void ButtonAttention_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    Outlook.ContactItem emailContact;
                    try
                    {
                        emailContact = InTouch.Contacts.FindContactFromEmailAddress(email.Sender.Address);
                    }
                    catch (Exception ex) 
                    { 
                        Log.Error(ex.Message, ex);
                        throw; 
                    }

                    if (emailContact is object)
                    {
                        InTouch.ShowInTouchSettings = true;
                        emailContact.Display(false);
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                    if (emailContact is object) { Marshal.ReleaseComObject(emailContact); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
            }
        }

        private void ButtonAddContactPersonal_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        Outlook.MAPIFolder contactsFolder = null;
                        Outlook.Items items = null;
                        Outlook.ContactItem contact = null;
                        try
                        {
                            InTouch.ShowInTouchSettings = true;
                            contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                            items = contactsFolder.Items;
                            contact = items.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;

                            Clipboard.SetDataObject(email.Sender.Name);

                            contact.FullName = email.Sender.Name;
                            contact.Email1Address = email.Sender.Address;

                            string data;
                            contact.UserProperties.Add("InTouchContact", Outlook.OlUserPropertyType.olText);
                            data = "|";
                            data += "|";
                            data += "0|";
                            data += "2|";
                            data += "2|";
                            data += "True|";

                            contact.UserProperties["InTouchContact"].Value = data;

                            contact.Save();
                            contact.Display(true);

                            lastEntryID = "";
                            //TODO Remove these.
                            Parallel.Invoke(() => { CheckEmailSender(); });
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            InTouch.ShowInTouchSettings = false;
                        }
                        finally
                        {
                            if (contact != null) Marshal.ReleaseComObject(contact);
                            if (items != null) Marshal.ReleaseComObject(items);
                            if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
                        }
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
            }
        }

        private void ButtonAddContactOther_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        Outlook.MAPIFolder contactsFolder = null;
                        Outlook.Items items = null;
                        Outlook.ContactItem contact = null;
                        try
                        {
                            InTouch.ShowInTouchSettings = true;
                            contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders["Others"];
                            items = contactsFolder.Items;
                            contact = items.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;

                            Clipboard.SetDataObject(email.Sender.Name);

                            contact.FullName = email.Sender.Name;
                            contact.Email1Address = email.Sender.Address;

                            string data;
                            contact.UserProperties.Add("InTouchContact", Outlook.OlUserPropertyType.olText);
                            data = "|";
                            data += "|";
                            data += "0|";
                            data += "2|";
                            data += "2|";
                            data += "True|";

                            contact.UserProperties["InTouchContact"].Value = data;

                            contact.Save();
                            contact.Display(true);

                            lastEntryID = "";
                            Parallel.Invoke(() => { CheckEmailSender(); });
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            InTouch.ShowInTouchSettings = false;
                        }
                        finally
                        {
                            if (contact != null) Marshal.ReleaseComObject(contact);
                            if (items != null) Marshal.ReleaseComObject(items);
                            if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
                        }
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
            }
        }

        private void ButtonAddContactJunk_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        Outlook.MAPIFolder contactsFolder = null;
                        Outlook.Items items = null;
                        Outlook.ContactItem contact = null;
                        try
                        {
                            InTouch.ShowInTouchSettings = true;
                            contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders["Junk Contacts"];
                            items = contactsFolder.Items;
                            contact = items.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;

                            Clipboard.SetDataObject(email.Sender.Name);

                            contact.FullName = email.Sender.Name;
                            contact.Email1Address = email.Sender.Address;

                            string data;
                            contact.UserProperties.Add("InTouchContact", Outlook.OlUserPropertyType.olText);
                            data = "|";
                            data += "|";
                            data += "3|";
                            data += "3|";
                            data += "0|";
                            data += "True|";
                            contact.UserProperties["InTouchContact"].Value = data;
                            
                            contact.Save();
                            contact.Display(true);

                            lastEntryID = "";
                            Parallel.Invoke(() => { CheckEmailSender(); });
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            InTouch.ShowInTouchSettings = false;
                        }
                        finally
                        {
                            if (contact != null) Marshal.ReleaseComObject(contact);
                            if (items != null) Marshal.ReleaseComObject(items);
                            if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
                        }
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
            }
        }
    }
}
