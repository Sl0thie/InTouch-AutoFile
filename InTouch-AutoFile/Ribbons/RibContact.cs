namespace InTouch_AutoFile
{
    using Microsoft.Office.Tools.Ribbon;

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    using Outlook = Microsoft.Office.Interop.Outlook;

    using Serilog;


    public partial class RibContact
    {
        private Outlook.Inspector inspector;

        private void RibContact_Load(object sender, RibbonUIEventArgs e)
        {
            inspector = Context as Outlook.Inspector;
            if (InTouch.ShowInTouchSettings)
            {
                inspector.SetCurrentFormPage("InTouch-AutoFile.ContactInTouchSettings");
                InTouch.ShowInTouchSettings = false;
            }
        }

        private void ButtonInTouchSettings_Click(object sender, RibbonControlEventArgs e)
        {
            inspector.SetCurrentFormPage("InTouch-AutoFile.ContactInTouchSettings");
        }
    }
}
