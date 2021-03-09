﻿
namespace InTouch_AutoFile
{
    partial class RibContact : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibContact()
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
            this.buttonInTouchSettings = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabContact";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabContact";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonInTouchSettings);
            this.group1.Label = "InTouch";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // buttonInTouchSettings
            // 
            this.buttonInTouchSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonInTouchSettings.Image = global::InTouch_AutoFile.Properties.Resources.options;
            this.buttonInTouchSettings.Label = "Settings";
            this.buttonInTouchSettings.Name = "buttonInTouchSettings";
            this.buttonInTouchSettings.ShowImage = true;
            this.buttonInTouchSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonInTouchSettings_Click);
            // 
            // RibContact
            // 
            this.Name = "RibContact";
            this.RibbonType = "Microsoft.Outlook.Contact";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibContact_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonInTouchSettings;
    }

    partial class ThisRibbonCollection
    {
        internal RibContact RibContact
        {
            get { return this.GetRibbon<RibContact>(); }
        }
    }
}
