using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InTouch_AutoFile
{
    partial class ContactInTouchSettings
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Contact)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("InTouch-AutoFile.ContactInTouchSettings")]
        public partial class ContactInTouchSettingsFactory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void ContactInTouchSettingsFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                
            }
        }

        #endregion

        InTouchContact contact;

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ContactInTouchSettings_FormRegionShowing(object sender, System.EventArgs e)
        {
            //change the colors on the FormRegion to suit the dark theme if needed.
            if ((BackColor.R == 38) && (BackColor.G == 38) && (BackColor.B == 38))
            {
                ForeColor = Color.White;
            }

            contact = new InTouchContact(this.OutlookItem as Outlook.ContactItem);

            LabelDeliveryPath.Text = contact.InboxPath;
            LabelReadPath.Text = contact.InboxPath;
            LabelSendPathValue.Text = contact.SentPath;

            switch (contact.DeliveryAction)
            {
                case EmailAction.None:
                    RadioButtonDeliveryNoAction.Checked = true;
                    break;

                case EmailAction.Delete:
                    RadioButtonDeliveryDelete.Checked = true;
                    break;

                case EmailAction.Move:
                    RadioButtonDeliveryFile.Checked = true;
                    break;
            }

            switch (contact.ReadAction)
            {
                case EmailAction.None:
                    RadioButtonReadNoAction.Checked = true;
                    break;

                case EmailAction.Delete:
                    RadioButtonReadDelete.Checked = true;
                    break;

                case EmailAction.Move:
                    RadioButtonReadFile.Checked = true;
                    break;
            }

            switch (contact.SentAction)
            {
                case EmailAction.None:
                    RadioButtonSendNoAction.Checked = true;
                    break;

                case EmailAction.Delete:
                    RadioButtonSendDelete.Checked = true;
                    break;

                case EmailAction.Move:
                    RadioButtonSendFile.Checked = true;
                    break;
            }
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ContactInTouchSettings_FormRegionClosed(object sender, System.EventArgs e)
        {
            contact.SaveAndDispose();
        }

        #region Control Events

        #region Delivery Action

        private void ButtonDeliveryPath_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder pickedFolder = outlookNameSpace.PickFolder();

            if (pickedFolder.FolderPath is object)
            {
                string backslash = @"\";
                string folderPath = pickedFolder.FolderPath;
                for (int i = 0; i < 4; i++)
                {
                    if (folderPath.IndexOf(backslash) >= 0)
                    {
                        folderPath = folderPath.Substring(folderPath.IndexOf(backslash) + 1);
                    }
                }

                LabelDeliveryPath.Text = folderPath;
                if (CheckBoxUseSamePath.Checked)
                {
                    LabelSendPathValue.Text = folderPath;
                }
            }
            else
            {
                LabelDeliveryPath.Text = "";
                if (CheckBoxUseSamePath.Checked)
                {
                    LabelSendPathValue.Text = "";
                }
            }

            //TODO Add check to see if this path is in the sent inbox folder.
            contact.InboxPath = LabelReadPath.Text;

            if (pickedFolder is object) { Marshal.ReleaseComObject(pickedFolder); }
            if (outlookNameSpace is object) { Marshal.ReleaseComObject(outlookNameSpace); }

        }

        private void RadioButtonDeliveryNoAction_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryNoAction.Checked)
            {
                contact.DeliveryAction = EmailAction.None;

                ButtonDeliveryPath.Visible = false;
                LabelDeliveryPathTitle.Visible = false;
                LabelDeliveryPath.Visible = false;
                CheckBoxUseSamePath.Visible = false;
            }
        }

        private void RadioButtonDeliveryDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryDelete.Checked)
            {
                contact.DeliveryAction = EmailAction.Delete;

                ButtonDeliveryPath.Visible = false;
                LabelDeliveryPathTitle.Visible = false;
                LabelDeliveryPath.Visible = false;
                CheckBoxUseSamePath.Visible = false;
            }
        }

        private void RadioButtonDeliveryFile_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryFile.Checked)
            {
                contact.DeliveryAction = EmailAction.Move;

                ButtonDeliveryPath.Visible = true;
                LabelDeliveryPathTitle.Visible = true;
                LabelDeliveryPath.Visible = true;
                CheckBoxUseSamePath.Visible = true;
            }
        }

        #endregion

        #region Read Action

        private void ButtonReadPath_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder pickedFolder = outlookNameSpace.PickFolder();

            if (pickedFolder.FolderPath is object)
            {
                string backslash = @"\";
                string folderPath = pickedFolder.FolderPath;
                for (int i = 0; i < 4; i++)
                {
                    if (folderPath.IndexOf(backslash) >= 0)
                    {
                        folderPath = folderPath.Substring(folderPath.IndexOf(backslash) + 1);
                    }
                }

                LabelReadPath.Text = folderPath;
            }
            else
            {
                LabelReadPath.Text = "";
            }

            //TODO Add check to see if this path is in the sent inbox folder.
            contact.InboxPath = LabelReadPath.Text;

            if (pickedFolder is object) { Marshal.ReleaseComObject(pickedFolder); }
            if (outlookNameSpace is object) { Marshal.ReleaseComObject(outlookNameSpace); }
        }

        private void RadioButtonReadNoAction_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadNoAction.Checked)
            {
                contact.ReadAction = EmailAction.None;

                ButtonReadPath.Visible = false;
                LabelPath.Visible = false;
                LabelReadPath.Visible = false;
            }
        }

        private void RadioButtonReadDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadDelete.Checked)
            {
                contact.ReadAction = EmailAction.Delete;

                ButtonReadPath.Visible = false;
                LabelPath.Visible = false;
                LabelReadPath.Visible = false;
            }
        }

        private void RadioButtonReadFile_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadFile.Checked)
            {
                contact.ReadAction = EmailAction.Move;

                ButtonReadPath.Visible = true;
                LabelPath.Visible = true;
                LabelReadPath.Visible = true;
            }
        }

        private void CheckBoxUseSamePath_CheckedChanged(object sender, EventArgs e)
        {
            //TODO Wire this up.
        }

        #endregion

        #region Send Action

        private void ButtonSendPath_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder pickedFolder = outlookNameSpace.PickFolder();

            if (pickedFolder.FolderPath is object)
            {
                string backslash = @"\";
                string folderPath = pickedFolder.FolderPath;
                for (int i = 0; i < 4; i++)
                {
                    if (folderPath.IndexOf(backslash) >= 0)
                    {
                        folderPath = folderPath.Substring(folderPath.IndexOf(backslash) + 1);
                    }
                }

                LabelSendPathValue.Text = folderPath;
            }
            else
            {
                LabelSendPathValue.Text = "";
            }

            //TODO Add check to see if this path is in the sent items folder.
            contact.SentPath = LabelSendPathValue.Text;

            if (pickedFolder is object) { Marshal.ReleaseComObject(pickedFolder); }
            if (outlookNameSpace is object) { Marshal.ReleaseComObject(outlookNameSpace); }
        }

        private void RadioButtonSendNoAction_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonSendNoAction.Checked)
            {
                contact.SentAction = EmailAction.None;

                ButtonSendPath.Visible = false;
                LabelSendPathTitle.Visible = false;
                LabelSendPathValue.Visible = false;
            }
        }

        private void RadioButtonSendDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonSendDelete.Checked)
            {
                contact.SentAction = EmailAction.Delete;

                ButtonSendPath.Visible = false;
                LabelSendPathTitle.Visible = false;
                LabelSendPathValue.Visible = false;
            }
        }

        private void RadioButtonSendFile_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonSendFile.Checked)
            {
                contact.SentAction = EmailAction.Move;

                ButtonSendPath.Visible = true;
                LabelSendPathTitle.Visible = true;
                LabelSendPathValue.Visible = true;
            }
        }

        #endregion

        #endregion
    }
}
