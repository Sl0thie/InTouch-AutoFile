using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace InTouch_AutoFile
{
    partial class MailItemHeader
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("InTouch-AutoFile.MailItemHeader")]
        public partial class MailItemHeaderFactory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void MailItemHeaderFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        Outlook.MailItem email;

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void MailItemHeader_FormRegionShowing(object sender, System.EventArgs e)
        {
            if (InTouch.DarkTheme)
            {
                RichText.BackColor = System.Drawing.Color.FromArgb(38, 38, 38);
                RichText.ForeColor = System.Drawing.Color.White;
            }

            email = this.OutlookItem as Outlook.MailItem;

            Outlook.PropertyAccessor mapiPropertyAccessor;
            //string propertyName = "http://schemas.microsoft.com/mapi/proptag/0x0065001F";
            string propertyName = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
            mapiPropertyAccessor = email.PropertyAccessor;
            string emailHeader = mapiPropertyAccessor.GetProperty(propertyName).ToString();
            if (mapiPropertyAccessor is object)
            {
                Marshal.ReleaseComObject(mapiPropertyAccessor);
            }

            RichText.Text = emailHeader;
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void MailItemHeader_FormRegionClosed(object sender, System.EventArgs e)
        {
            if (email is object) Marshal.ReleaseComObject(email);
        }

        private void MailItemHeader_Resize(object sender, EventArgs e)
        {
            RichText.Width = this.Width;
            RichText.Height = this.Height;           
        }
    }
}
