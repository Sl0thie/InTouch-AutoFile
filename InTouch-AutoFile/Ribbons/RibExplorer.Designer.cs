
namespace InTouch_AutoFile
{
    partial class RibExplorer : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibExplorer()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.buttonContact = this.Factory.CreateRibbonButton();
            this.buttonAddContactPersonal = this.Factory.CreateRibbonButton();
            this.buttonAddContactOther = this.Factory.CreateRibbonButton();
            this.buttonAddContactJunk = this.Factory.CreateRibbonButton();
            this.buttonAttention = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonContact);
            this.group1.Items.Add(this.buttonAddContactPersonal);
            this.group1.Items.Add(this.buttonAddContactOther);
            this.group1.Items.Add(this.buttonAddContactJunk);
            this.group1.Items.Add(this.buttonAttention);
            this.group1.Label = "InTouch";
            this.group1.Name = "group1";
            // 
            // buttonContact
            // 
            this.buttonContact.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonContact.Image = global::InTouch_AutoFile.Properties.Resources.contact;
            this.buttonContact.Label = "Contact";
            this.buttonContact.Name = "buttonContact";
            this.buttonContact.ShowImage = true;
            this.buttonContact.Visible = false;
            this.buttonContact.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonContact_Click);
            // 
            // buttonAddContactPersonal
            // 
            this.buttonAddContactPersonal.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddContactPersonal.Image = global::InTouch_AutoFile.Properties.Resources.addcontact;
            this.buttonAddContactPersonal.Label = "Add to Personal Contacts";
            this.buttonAddContactPersonal.Name = "buttonAddContactPersonal";
            this.buttonAddContactPersonal.ShowImage = true;
            this.buttonAddContactPersonal.Visible = false;
            this.buttonAddContactPersonal.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAddContactPersonal_Click);
            // 
            // buttonAddContactOther
            // 
            this.buttonAddContactOther.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddContactOther.Image = global::InTouch_AutoFile.Properties.Resources.addcontact;
            this.buttonAddContactOther.Label = "Add to Other Contacts";
            this.buttonAddContactOther.Name = "buttonAddContactOther";
            this.buttonAddContactOther.ShowImage = true;
            this.buttonAddContactOther.Visible = false;
            this.buttonAddContactOther.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAddContactOther_Click);
            // 
            // buttonAddContactJunk
            // 
            this.buttonAddContactJunk.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAddContactJunk.Image = global::InTouch_AutoFile.Properties.Resources.addcontact;
            this.buttonAddContactJunk.Label = "Junk Mail";
            this.buttonAddContactJunk.Name = "buttonAddContactJunk";
            this.buttonAddContactJunk.ShowImage = true;
            this.buttonAddContactJunk.Visible = false;
            this.buttonAddContactJunk.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAddContactJunk_Click);
            // 
            // buttonAttention
            // 
            this.buttonAttention.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAttention.Image = global::InTouch_AutoFile.Properties.Resources.repath;
            this.buttonAttention.Label = "Attention Required";
            this.buttonAttention.Name = "buttonAttention";
            this.buttonAttention.ShowImage = true;
            this.buttonAttention.Visible = false;
            this.buttonAttention.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonAttention_Click);
            // 
            // RibExplorer
            // 
            this.Name = "RibExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibExplorer_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddContactPersonal;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAttention;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonContact;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddContactOther;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddContactJunk;
    }

    partial class ThisRibbonCollection
    {
        internal RibExplorer RibExplorer
        {
            get { return this.GetRibbon<RibExplorer>(); }
        }
    }
}
