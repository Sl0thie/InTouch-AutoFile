namespace InTouch_AutoFile
{
    using Microsoft.Office.Tools.Ribbon;
    using System;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using System.Drawing;
    using System.Windows.Forms;
    using System.Threading.Tasks;
    using Serilog;

    //TODO Redo the task to include cancellation so they don't run into each other.
    /// <summary>
    /// A Ribbon extension for the Outlook Explorer Menu to provide buttons for InTouch.
    /// </summary>
    public partial class RibExplorer
    {
        private Outlook.Explorer explorer; // The current explorer object.

        private string lastEntryID = "";

        /// <summary>
        /// Constructor for the explorer ribbon.
        /// This ribbon show when emails are selected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RibExplorer_Load(object sender, RibbonUIEventArgs e)
        {
            // Hook into the current explorer object and wire up the selection change event.
            explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            explorer.SelectionChange += Explorer_SelectionChange;

            // Fire for first event that is missed during startup.
            Task.Factory.StartNew(() => CheckEmailSender());
        }

        /// <summary>
        /// Event handler for when the item is changed in outlook. 
        /// If it is a email then manage the buttons on the ribbon.
        /// </summary>
        private void Explorer_SelectionChange()
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                // Check the first object selected.
                dynamic selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];

                if (selectedObject.EntryID != lastEntryID)
                {
                    lastEntryID = selectedObject.EntryID;
                    ClearButtons();

                    // If it's a mail object then manage based on the contacts details.
                    if (selectedObject is Outlook.MailItem)
                    {
                        Task.Factory.StartNew(() => CheckEmailSender());
                    }
                    else if (selectedObject is Outlook.TaskItem)
                    {
                        lastEntryID = "";
                    }
                    else if (selectedObject is Outlook.ContactItem)
                    {
                        lastEntryID = "";
                    }
                    else if (selectedObject is Outlook.AppointmentItem)
                    {
                        lastEntryID = "";
                    }
                }

                // Release Outlook objects.
                if (selectedObject is object)
                {
                    Marshal.ReleaseComObject(selectedObject);
                }
            }
            else
            {
                ClearButtons();
            }
        }

        /// <summary>
        /// ClearButtons method makes all the ribbon buttons invisible.
        /// </summary>
        private static void ClearButtons()
        {
            Globals.Ribbons.RibExplorer.buttonContact.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAddContactPersonal.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAddContactOther.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAddContactJunk.Visible = false;
            Globals.Ribbons.RibExplorer.buttonAttention.Visible = false;

            //TODO re-check if this is still required.
            Application.DoEvents();
        }

        /// <summary>
        /// CheckEmailSender method is called when the selected email changes. 
        /// It first finds the contact that sent the email.
        /// It then sets the ribbon to suit that contact.
        /// </summary>
        private static void CheckEmailSender()
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
                    // Try to find contact from email address.
                    InTouchContact emailContact = null;
                    try
                    {
                        Outlook.ContactItem contact = Contacts.FindContactFromEmailAddress(email.Sender.Address);
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
                        // Make the Contact Button visible and add the image and name to the button.
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

                        // Check if the contact details are valid.
                        if (!emailContact.CheckDetails())
                        {
                            Globals.Ribbons.RibExplorer.buttonAttention.Visible = true;
                        }

                        emailContact.SaveAndDispose();
                    }
                    else
                    {
                        // As the contact was not found, make the add contact button visible.
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
            }

            // Release Outlook objects.
            if (email is object)
            {
                Marshal.ReleaseComObject(email);
            }

            if (selection is object)
            {
                Marshal.ReleaseComObject(selection);
            }
        }

        /// <summary>
        /// ButtonContact_Click method handles the Contact button's click event.
        /// It finds the contact from the email and then displays the contact.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonContact_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    // Find contact from sender email address.
                    Outlook.ContactItem emailContact;
                    try
                    {
                        emailContact = Contacts.FindContactFromEmailAddress(email.Sender.Address);
                    }
                    catch (Exception ex) 
                    { 
                        Log.Error(ex.Message,ex);
                        throw;
                    }

                    // Display contact.
                    if (emailContact is object)
                    {
                        emailContact.Display(false);
                    }

                    // Release Outlook objects.
                    if (email is object)
                    {
                        Marshal.ReleaseComObject(email);
                    }

                    if (emailContact is object)
                    {
                        Marshal.ReleaseComObject(emailContact);
                    }
                }

                // Release Outlook objects.
                if (selectedObject is object)
                {
                    Marshal.ReleaseComObject(selectedObject);
                }
            }
        }

        /// <summary>
        /// ButtonAttention_Click method handles the Attention button's click event.
        /// It finds the contact from the email and then displays the contact's InTouch form region.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAttention_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    // Find contact from sender email address.
                    Outlook.ContactItem emailContact;
                    try
                    {
                        emailContact = Contacts.FindContactFromEmailAddress(email.Sender.Address);
                    }
                    catch (Exception ex) 
                    { 
                        Log.Error(ex.Message, ex);
                        throw; 
                    }

                    // Display contact, showing the InTouch form region.
                    if (emailContact is object)
                    {
                        InTouch.ShowInTouchSettings = true;
                        emailContact.Display(false);
                    }

                    // Release Outlook objects.
                    if (email is object)
                    {
                        Marshal.ReleaseComObject(email);
                    }

                    if (emailContact is object)
                    {
                        Marshal.ReleaseComObject(emailContact);
                    }
                }

                // Release Outlook objects.
                if (selectedObject is object)
                {
                    Marshal.ReleaseComObject(selectedObject);
                }
            }
        }

        /// <summary>
        /// ButtonAddContactPersonal_Click method handles the Add Contact button's click event.
        /// It creates a contact from the emails sender and adds it to the personal contact store.
        /// It then displays the contact's InTouch form region.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAddContactPersonal_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        Outlook.MAPIFolder contactsFolder = null;
                        Outlook.Items items = null;
                        Outlook.ContactItem contact = null;
                        try
                        {
                            // Create new contact from the email's sender.
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
                            Contacts.AddContactToEmailLookup(contact, (Outlook.Folder)contactsFolder);
                            lastEntryID = "";
                            Parallel.Invoke(() => CheckEmailSender());
                            InTouch.TaskManager.EnqueueInboxTask();
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            InTouch.ShowInTouchSettings = false;
                        }
                        finally
                        {
                            // Release Outlook objects.
                            if (contact != null)
                            {
                                Marshal.ReleaseComObject(contact);
                            }

                            if (items != null)
                            {
                                Marshal.ReleaseComObject(items);
                            }

                            if (contactsFolder != null)
                            {
                                Marshal.ReleaseComObject(contactsFolder);
                            }
                        }
                    }

                    // Release Outlook objects.
                    if (email is object)
                    {
                        Marshal.ReleaseComObject(email);
                    }
                }

                // Release Outlook objects.
                if (selectedObject is object)
                {
                    Marshal.ReleaseComObject(selectedObject);
                }
            }
        }

        /// <summary>
        /// ButtonAddContactOther_Click method handles the Add Contact Other button's click event.
        /// It creates a contact from the emails sender and adds it to the other contacts store.
        /// It then displays the contact's InTouch form region.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAddContactOther_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        Outlook.MAPIFolder contactsFolder = null;
                        Outlook.Items items = null;
                        Outlook.ContactItem contact = null;
                        try
                        {
                            // Create new contact from the email's sender.
                            InTouch.ShowInTouchSettings = true;
                            contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders[InTouch.OtherFolderName];
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

                            // Save and display the contact to the user.
                            contact.Save();
                            contact.Display(true);
                            Contacts.AddContactToEmailLookup(contact, (Outlook.Folder)contactsFolder);
                            lastEntryID = "";
                            Parallel.Invoke(() => CheckEmailSender());
                            InTouch.TaskManager.EnqueueInboxTask();
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            InTouch.ShowInTouchSettings = false;
                        }
                        finally
                        {
                            // Release Outlook objects.
                            if (contact != null)
                            {
                                Marshal.ReleaseComObject(contact);
                            }

                            if (items != null)
                            {
                                Marshal.ReleaseComObject(items);
                            }

                            if (contactsFolder != null)
                            {
                                Marshal.ReleaseComObject(contactsFolder);
                            }
                        }
                    }

                    // Release Outlook objects.
                    if (email is object)
                    {
                        Marshal.ReleaseComObject(email);
                    }
                }

                // Release Outlook objects.
                if (selectedObject is object)
                {
                    Marshal.ReleaseComObject(selectedObject);
                }
            }
        }

        /// <summary>
        /// ButtonAddContactJunk_Click method handles the Add Contact Junk button's click event.
        /// It creates a contact from the emails sender and adds it to the junk contacts store.
        /// It then displays the contact's InTouch form region.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonAddContactJunk_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        Outlook.MAPIFolder contactsFolder = null;
                        Outlook.Items items = null;
                        Outlook.ContactItem contact = null;
                        try
                        {
                            // Create new contact from the email's sender.
                            InTouch.ShowInTouchSettings = true;
                            contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders[InTouch.JunkFolderName];
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
                            //contact.Display(true);

                            // Move to junk email folder.
                            email.Move(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk));

                            lastEntryID = "";
                            Parallel.Invoke(() => CheckEmailSender());
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex.Message, ex);
                            InTouch.ShowInTouchSettings = false;
                        }
                        finally
                        {
                            // Release Outlook objects.
                            if (contact != null)
                            {
                                Marshal.ReleaseComObject(contact);
                            }

                            if (items != null)
                            {
                                Marshal.ReleaseComObject(items);
                            }

                            if (contactsFolder != null)
                            {
                                Marshal.ReleaseComObject(contactsFolder);
                            }
                        }
                    }

                    // Release Outlook objects.
                    if (email is object)
                    {
                        Marshal.ReleaseComObject(email);
                    }
                }

                // Release Outlook objects.
                if (selectedObject is object)
                {
                    Marshal.ReleaseComObject(selectedObject);
                }
            }
        }
    }
}
