namespace InTouch_AutoFile
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using System.Runtime.InteropServices;
    using Outlook = Microsoft.Office.Interop.Outlook;

    [ComVisible(true)]
    public sealed partial class UCPropertPage : UserControl, Outlook.PropertyPage
    {

        #region Required

        private const int captionDispID = -518;
        private bool isDirty = false;
        private Outlook.PropertyPageSite ppSite;

        public UCPropertPage()
        {
            InitializeComponent();
        }

        void Outlook.PropertyPage.Apply()
        {
            Properties.Settings.Default.TaskInbox = CheckBoxTaskInbox.Checked;
            Properties.Settings.Default.TaskSent = CheckBoxTaskSent.Checked;
            Properties.Settings.Default.TaskDuplicates = CheckBoxTaskDuplicates.Checked;
            Properties.Settings.Default.TaskEmailRouting = CheckBoxTaskRouting.Checked;
            Properties.Settings.Default.Save();
        }

        bool Outlook.PropertyPage.Dirty
        {
            get
            {
                return isDirty;
            }
        }

        void Outlook.PropertyPage.GetPageInfo(ref string helpFile, ref int helpContext)
        {

        }

        [DispId(captionDispID)]
        public string PageCaption
        {
            get
            {
                return "InTouch";
            }
        }

        private void UCPropertPage_Load(object sender, EventArgs e)
        {
            //Required to make 'Apply' button work.
            Type myType = typeof(object);
            string assembly = System.Text.RegularExpressions.Regex.Replace(myType.Assembly.CodeBase, "mscorlib.dll", "System.Windows.Forms.dll");
            assembly = System.Text.RegularExpressions.Regex.Replace(assembly, "file:///", "");
            assembly = System.Reflection.AssemblyName.GetAssemblyName(assembly).FullName;
            Type unmanaged = Type.GetType(System.Reflection.Assembly.CreateQualifiedName(assembly, "System.Windows.Forms.UnsafeNativeMethods"));
            Type oleObj = unmanaged.GetNestedType("IOleObject");
            System.Reflection.MethodInfo mi = oleObj.GetMethod("GetClientSite");
            object myppSite = mi.Invoke(this, null);
            ppSite = (Outlook.PropertyPageSite)myppSite;

            CheckBoxTaskInbox.Checked = Properties.Settings.Default.TaskInbox;
            CheckBoxTaskSent.Checked = Properties.Settings.Default.TaskSent;
            CheckBoxTaskDuplicates.Checked = Properties.Settings.Default.TaskDuplicates;
            CheckBoxTaskRouting.Checked = Properties.Settings.Default.TaskEmailRouting;
        }

        #endregion

        private void CheckBoxShowTasksButton_CheckedChanged(object sender, EventArgs e)
        {
            isDirty = true;
            ppSite.OnStatusChange();
        }

        private void CheckBoxTaskInbox_CheckedChanged(object sender, EventArgs e)
        {
            isDirty = true;
            ppSite.OnStatusChange();
        }

        private void CheckBoxTaskSent_CheckedChanged(object sender, EventArgs e)
        {
            isDirty = true;
            ppSite.OnStatusChange();
        }

        private void CheckBoxTaskDuplicates_CheckedChanged(object sender, EventArgs e)
        {
            isDirty = true;
            ppSite.OnStatusChange();
        }

        private void CheckBoxTaskRouting_CheckedChanged(object sender, EventArgs e)
        {
            isDirty = true;
            ppSite.OnStatusChange();
        }
    }
}
