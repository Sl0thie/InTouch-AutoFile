using Microsoft.Office.Tools.Ribbon;
using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;

namespace InTouch_AutoFile
{
    /// <summary>
    /// A Ribbon extention for the Outlook Explorer Menu to provide buttons for InTouch.
    /// </summary>
    public partial class RibExplorer
    {
        private Outlook.Explorer explorer; //The current explorer object.

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

            if (Properties.Settings.Default.ShowTasksButton)
            {
                ButtonTasks.Visible = true;
            }
        }

        /// <summary>
        /// Event handler for when the item is changed in outlook. 
        /// If it is a email then manage the buttons on the ribbon.
        /// </summary>
        private void Explorer_SelectionChange()
        {
            if (Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                //Clear all the buttons.
                Globals.Ribbons.RibExplorer.buttonContact.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAddContact.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAttention.Visible = false;

                //check the first object selected.
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];

                //If it's a mail object then manage based on the contacts details.
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email.Sender is object)
                    {
                        //Try to find contact from email adddress.
                        InTouchContact emailContact = null;
                        try
                        {
                            Outlook.ContactItem contact = InTouch.Contacts.FindContactFromEmailAddress(email.Sender.Address);
                            if(contact is object)
                            {
                                emailContact = new InTouchContact(contact);
                            }
                        }
                        catch (Exception ex)
                        {
                            Op.LogError(ex);
                        }

                        if (emailContact is object)
                        {
                            //Make the Contact Button visable and add the image and name to the button.
                            Globals.Ribbons.RibExplorer.buttonContact.Visible = true;
                            
                            if(emailContact.FullName is object)
                            {
                                Globals.Ribbons.RibExplorer.buttonContact.Label = emailContact.FullName;
                            }
                            else
                            {
                                Globals.Ribbons.RibExplorer.buttonContact.Label = "";
                            }

                            if (emailContact.HasPicture)
                            {
                                Globals.Ribbons.RibExplorer.buttonContact.Image = Image.FromFile(GetContactPicturePath(emailContact));
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
                            Globals.Ribbons.RibExplorer.buttonAddContact.Visible = true;
                        }
                    }
                    else
                    {
                        Globals.Ribbons.RibExplorer.buttonAddContact.Visible = true;
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                }
                else if (selectedObject is Outlook.TaskItem) { }
                else if (selectedObject is Outlook.ContactItem) { }
                else if (selectedObject is Outlook.AppointmentItem) { }
                if (selectedObject is object) Marshal.ReleaseComObject(selectedObject);
            }
            else
            {
                //No object is selected so remove all buttons.
                Globals.Ribbons.RibExplorer.buttonAddContact.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAddContact.Visible = false;
                Globals.Ribbons.RibExplorer.buttonAttention.Visible = false;
            }
        }

        /// <summary>
        /// Gets the path to the contacts picture to be used in the ribbon's button.
        /// </summary>
        /// <param name="contact"></param>
        /// <returns></returns>
        public static string GetContactPicturePath(Outlook._ContactItem contact)
        {
            if(contact is object)
            {
                string picturePath = "";
                if (contact.HasPicture)
                {
                    foreach (Outlook.Attachment att in contact.Attachments)
                    {
                        if (att.DisplayName == "ContactPicture.jpg")
                        {
                            try
                            {
                                picturePath = System.IO.Path.GetDirectoryName(System.IO.Path.GetTempPath()) + "\\Contact_" + contact.EntryID + ".jpg";
                                if (!System.IO.File.Exists(picturePath))
                                {
                                    att.SaveAsFile(picturePath);
                                }
                            }
                            catch (COMException)
                            {
                                picturePath = "";
                            }
                            catch (Exception ex)
                            {
                                Op.LogError(ex);
                                picturePath = "";
                                throw;
                            }
                        }
                    }
                }
                return picturePath;
            }
            else
            {
                return null;
            }
        }

        private void ButtonAddContact_Click(object sender, RibbonControlEventArgs e)
        {
            if(Globals.ThisAddIn.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedObject is Outlook.MailItem email)
                {
                    if (email is object)
                    {
                        using (FormInTouchNewContact newContactForm = new FormInTouchNewContact())
                        {
                            newContactForm.Email = email;
                            newContactForm.ShowDialog();

                            Application.DoEvents();
                            Thread.Sleep(1000);

                            if (Op.ContactCreatedEmail is object) 
                            {
                                Outlook.ContactItem contact = InTouch.Contacts.FindContactFromEmailAddress(Op.ContactCreatedEmail);
                                if(contact is object)
                                {
                                    contact.Display();
                                }      
                            }
                        }
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
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
                        Op.NextFormRegion = ContactFormRegion.InTouchSettings;
                        emailContact.Display(false);
                    }
                    if (email is object) { Marshal.ReleaseComObject(email); }
                    if (emailContact is object) { Marshal.ReleaseComObject(emailContact); }
                }
                if (selectedObject is object) { Marshal.ReleaseComObject(selectedObject); }
            }
        }
    }
}
