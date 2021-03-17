using Microsoft.Office.Tools.Ribbon;
using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;

namespace InTouch_AutoFile
{
    /// <summary>
    /// A Ribbon extention for the Outlook Explorer Menu to provide buttons for InTouch.
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
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];

                //If it's a mail object then manage based on the contacts details.
                if (selectedObject is Outlook.MailItem)
                {
                    Parallel.Invoke(() => { CheckEmailSender(); });
                }
                else if (selectedObject is Outlook.TaskItem) { }
                else if (selectedObject is Outlook.ContactItem) { }
                else if (selectedObject is Outlook.AppointmentItem) { }
                if (selectedObject is object) Marshal.ReleaseComObject(selectedObject);
            }
            else
            {
                //Clear all the buttons.
                Globals.Ribbons.RibExplorer.buttonContact.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAddContactPersonal.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAddContactOther.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAddContactJunk.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAttention.Visible = false;
            }
        }

        private void CheckEmailSender()
        {
            Outlook.Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
            Outlook.MailItem email = selection[1] as Outlook.MailItem;

            if (email is object)
            {
                
                if(email.EntryID != lastEntryID)
                {
                    lastEntryID = email.EntryID;

                    //Clear all the buttons.
                    Globals.Ribbons.RibExplorer.buttonContact.Visible = false;
                    Globals.Ribbons.RibExplorer.buttonAddContactPersonal.Visible = false;
                    Globals.Ribbons.RibExplorer.buttonAddContactOther.Visible = false;
                    Globals.Ribbons.RibExplorer.buttonAddContactJunk.Visible = false;
                    Globals.Ribbons.RibExplorer.buttonAttention.Visible = false;

                    if (email.Sender is object)
                    {
                        //Try to find contact from email adddress.
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
                            Op.LogError(ex);
                            throw;
                        }

                        if (emailContact is object)
                        {
                            //Make the Contact Button visable and add the image and name to the button.
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
                    if (email is object) { Marshal.ReleaseComObject(email); }
                }
            }
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
                        Op.LogError(ex); 
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
                        Op.LogError(ex);
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




                            contact.Save();
                            contact.Display(true);
                        }
                        catch (Exception ex)
                        {
                            Op.LogError(ex);
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
                            contactsFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Folders["Other Contacts"];
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

                            contact.Save();
                            contact.Display(true);
                        }
                        catch (Exception ex)
                        {
                            Op.LogError(ex);
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

                            contact.Save();
                            contact.Display(true);
                        }
                        catch (Exception ex)
                        {
                            Op.LogError(ex);
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
