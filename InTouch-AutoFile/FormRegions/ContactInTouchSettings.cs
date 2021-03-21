using System;
using System.Drawing;
using System.Runtime.InteropServices;
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

                ButtonDeliveryPath.ForeColor = Color.Black;
                ButtonReadPath.ForeColor = Color.Black;
                ButtonSendPath.ForeColor = Color.Black;
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

                case EmailAction.Junk:
                    RadioButtonDeliveryJunk.Checked = true;
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

                case EmailAction.Junk:
                    RadioButtonReadJunk.Checked = true;
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

            CheckBoxUseSamePathDelivery.Checked = contact.SamePath;

            AdjustForm();
        }

        /// <summary>
        /// Adjust the Controls on the Form Region
        /// </summary>
        /// <remarks>
        /// This is currently going though trial and error to see what suits best.
        /// The issue of Inbox/SentItem/Junk folders has come up
        /// where you can't choose the junk folder as a place to file mail.
        /// and the sent folder is usally the same as the Inbox folder.
        /// </remarks>
        private void AdjustForm()
        {
            bool readAction = false;
            //Read actions are dependant on delivery actions because they happen after.
            if (RadioButtonDeliveryNoAction.Checked)
            {
                RadioButtonReadNoAction.Enabled = true;
                
                RadioButtonReadDelete.Enabled = true;
                RadioButtonReadFile.Enabled = true;

                ButtonDeliveryPath.Visible = false;
                LabelDeliveryPathTitle.Visible = false;
                LabelDeliveryPath.Visible = false;
                CheckBoxUseSamePathDelivery.Visible = false;
                readAction = true;
            }
            else if (RadioButtonDeliveryDelete.Checked)
            {
                //Force the read action to be the same.
                RadioButtonReadDelete.Checked = true;

                RadioButtonReadNoAction.Enabled = false;
                RadioButtonReadDelete.Enabled = false;
                RadioButtonReadFile.Enabled = false;

                ButtonDeliveryPath.Visible = false;
                LabelDeliveryPathTitle.Visible = false;
                LabelDeliveryPath.Visible = false;
                CheckBoxUseSamePathDelivery.Visible = false;

                ButtonReadPath.Visible = false;
                LabelPath.Visible = false;
                LabelReadPath.Visible = false;

            }
            else if (RadioButtonDeliveryFile.Checked)
            {
                //Force the read action to be the same.
                RadioButtonReadFile.Checked = true;

                RadioButtonReadNoAction.Enabled = false;
                RadioButtonReadDelete.Enabled = false;
                RadioButtonReadFile.Enabled = false;

                ButtonDeliveryPath.Visible = true;
                LabelDeliveryPathTitle.Visible = true;
                LabelDeliveryPath.Visible = true;
                CheckBoxUseSamePathDelivery.Visible = true;

                ButtonReadPath.Visible = false;
                LabelPath.Visible = false;
                LabelReadPath.Visible = false;

            }
            else
            {
                //Force the read action to be the same.
                RadioButtonReadJunk.Checked = true;

                RadioButtonReadNoAction.Enabled = false;
                RadioButtonReadDelete.Enabled = false;
                RadioButtonReadFile.Enabled = false;

                ButtonDeliveryPath.Visible = true;
                LabelDeliveryPathTitle.Visible = true;
                LabelDeliveryPath.Visible = true;
                CheckBoxUseSamePathDelivery.Visible = true;

                ButtonReadPath.Visible = false;
                LabelPath.Visible = false;
                LabelReadPath.Visible = false;
            }


            if (readAction)
            {
                if (RadioButtonReadNoAction.Checked)
                {
                    ButtonReadPath.Visible = false;
                    LabelPath.Visible = false;
                    LabelReadPath.Visible = false;

                }
                else if (RadioButtonReadDelete.Checked)
                {
                    ButtonReadPath.Visible = false;
                    LabelPath.Visible = false;
                    LabelReadPath.Visible = false;

                }
                else
                {
                    ButtonReadPath.Visible = true;
                    LabelPath.Visible = true;
                    LabelReadPath.Visible = true;

                }
            }

            if (RadioButtonSendNoAction.Checked)
            {
                ButtonSendPath.Visible = false;
                LabelSendPathTitle.Visible = false;
                LabelSendPathValue.Visible = false;

            }
            else if (RadioButtonSendDelete.Checked)
            {
                ButtonSendPath.Visible = false;
                LabelSendPathTitle.Visible = false;
                LabelSendPathValue.Visible = false;
            }
            else
            {
                ButtonSendPath.Visible = true;
                LabelSendPathTitle.Visible = true;
                LabelSendPathValue.Visible = true;

                if (CheckBoxUseSamePathDelivery.Checked) //Only check the first checkbox as they are tied together.
                {
                    ButtonSendPath.Visible = false;
                    LabelSendPathTitle.Visible = false;
                    LabelSendPathValue.Visible = false;
                }
                else
                {
                    ButtonSendPath.Visible = true;
                    LabelSendPathTitle.Visible = true;
                    LabelSendPathValue.Visible = true;
                }
            }
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void ContactInTouchSettings_FormRegionClosed(object sender, System.EventArgs e)
        {
            contact.SaveAndDispose();
        }

        private void GetSamePath()
        {
            string[] folders = contact.InboxPath.Split('\\');

            if(folders[1] == "Inbox")
            {
                string newPath = folders[0] + "\\Sent Items";
                for (int i = 2; i <= folders.GetUpperBound(0); i++)
                {
                    newPath += "\\" + folders[i];
                }

                contact.SentPath = newPath;
                InTouch.CreatePath(contact.SentPath);
                LabelSendPathValue.Text = contact.SentPath;
            }
        }

        #region Control Events

        #region Delivery Action

        private void ButtonDeliveryPath_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder pickedFolder = outlookNameSpace.PickFolder();
            string folderPath;
            if (pickedFolder.FolderPath is object)
            {
                folderPath = pickedFolder.FolderPath;
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
            }
            else
            {
                folderPath = "";
            }

            LabelDeliveryPath.Text = folderPath;
            LabelReadPath.Text = folderPath;
            contact.InboxPath = folderPath;


            if (contact.SamePath)
            {
                GetSamePath();
            }

            if (pickedFolder is object) { Marshal.ReleaseComObject(pickedFolder); }
            if (outlookNameSpace is object) { Marshal.ReleaseComObject(outlookNameSpace); }
        }

        private void RadioButtonDeliveryNoAction_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryNoAction.Checked)
            {
                contact.DeliveryAction = EmailAction.None;
                AdjustForm();
            }
        }

        private void RadioButtonDeliveryDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryDelete.Checked)
            {
                contact.DeliveryAction = EmailAction.Delete;
                AdjustForm();
            }
        }

        private void RadioButtonDeliveryFile_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryFile.Checked)
            {
                contact.DeliveryAction = EmailAction.Move;
                AdjustForm();
            }
        }

        private void RadioButtonDeliveryJunk_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonDeliveryJunk.Checked)
            {
                contact.DeliveryAction = EmailAction.Junk;
                AdjustForm();
            }
        }

        private void CheckBoxUseSamePathDelivery_CheckedChanged(object sender, EventArgs e)
        {
            contact.SamePath = CheckBoxUseSamePathDelivery.Checked;

            AdjustForm();
        }

        #endregion

        #region Read Action

        private void ButtonReadPath_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder pickedFolder = outlookNameSpace.PickFolder();
            string folderPath;
            if (pickedFolder.FolderPath is object)
            {
                folderPath = pickedFolder.FolderPath;
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
            }
            else
            {
                folderPath = "";
            }

            LabelReadPath.Text = folderPath;
            LabelDeliveryPath.Text = folderPath;
            contact.InboxPath = folderPath;

            if (contact.SamePath)
            {
                GetSamePath();
            }

            if (pickedFolder is object) { Marshal.ReleaseComObject(pickedFolder); }
            if (outlookNameSpace is object) { Marshal.ReleaseComObject(outlookNameSpace); }
        }

        private void RadioButtonReadNoAction_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadNoAction.Checked)
            {
                contact.ReadAction = EmailAction.None;
                AdjustForm();
            }
        }

        private void RadioButtonReadDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadDelete.Checked)
            {
                contact.ReadAction = EmailAction.Delete;
                AdjustForm();
            }
        }

        private void RadioButtonReadFile_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadFile.Checked)
            {
                contact.ReadAction = EmailAction.Move;
                AdjustForm();
            }
        }

        private void RadioButtonReadJunk_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonReadJunk.Checked)
            {
                contact.ReadAction = EmailAction.Junk;
                AdjustForm();
            }
        }

        #endregion

        #region Send Action

        private void ButtonSendPath_Click(object sender, EventArgs e)
        {
            Outlook.NameSpace outlookNameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder pickedFolder = outlookNameSpace.PickFolder();
            string folderPath;
            if (pickedFolder.FolderPath is object)
            {
                folderPath = pickedFolder.FolderPath;
                if (folderPath.StartsWith(@"\\"))
                {
                    folderPath = folderPath.Remove(0, 2);
                }
                Log.Message("FolderPath : " + folderPath);
            }
            else
            {
                folderPath = "";
            }

            LabelSendPathValue.Text = folderPath;
            contact.SentPath = folderPath;

            if (pickedFolder is object) { Marshal.ReleaseComObject(pickedFolder); }
            if (outlookNameSpace is object) { Marshal.ReleaseComObject(outlookNameSpace); }
        }

        private void RadioButtonSendNoAction_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonSendNoAction.Checked)
            {
                contact.SentAction = EmailAction.None;
                AdjustForm();
            }
        }

        private void RadioButtonSendDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonSendDelete.Checked)
            {
                contact.SentAction = EmailAction.Delete;
                AdjustForm();
            }
        }

        private void RadioButtonSendFile_CheckedChanged(object sender, EventArgs e)
        {
            if (RadioButtonSendFile.Checked)
            {
                contact.SentAction = EmailAction.Move;
                AdjustForm();
            }
        }




        #endregion

        #endregion

        
    }
}
