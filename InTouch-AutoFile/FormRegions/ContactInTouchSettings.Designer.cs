﻿
namespace InTouch_AutoFile
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class ContactInTouchSettings : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public ContactInTouchSettings(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.InitializeComponent();
        }

        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Form Region Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private static void InitializeManifest(Microsoft.Office.Tools.Outlook.FormRegionManifest manifest, Microsoft.Office.Tools.Outlook.Factory factory)
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ContactInTouchSettings));
            manifest.FormRegionName = "InTouch Settings";
            manifest.Icons.Default = ((System.Drawing.Icon)(resources.GetObject("ContactInTouchSettings.Manifest.Icons.Default")));
            manifest.Icons.Page = ((System.Drawing.Image)(resources.GetObject("ContactInTouchSettings.Manifest.Icons.Page")));
            manifest.Icons.Window = ((System.Drawing.Icon)(resources.GetObject("ContactInTouchSettings.Manifest.Icons.Window")));
            manifest.ShowReadingPane = false;

        }

        #endregion

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("InBox");
            this.PanelSend = new System.Windows.Forms.Panel();
            this.RadioButtonSendDelete = new System.Windows.Forms.RadioButton();
            this.RadioButtonSendFile = new System.Windows.Forms.RadioButton();
            this.RadioButtonSendNoAction = new System.Windows.Forms.RadioButton();
            this.PanelRead = new System.Windows.Forms.Panel();
            this.RadioButtonReadJunk = new System.Windows.Forms.RadioButton();
            this.RadioButtonReadDelete = new System.Windows.Forms.RadioButton();
            this.RadioButtonReadFile = new System.Windows.Forms.RadioButton();
            this.RadioButtonReadNoAction = new System.Windows.Forms.RadioButton();
            this.PanelDelivery = new System.Windows.Forms.Panel();
            this.RadioButtonDeliveryJunk = new System.Windows.Forms.RadioButton();
            this.RadioButtonDeliveryDelete = new System.Windows.Forms.RadioButton();
            this.RadioButtonDeliveryFile = new System.Windows.Forms.RadioButton();
            this.RadioButtonDeliveryNoAction = new System.Windows.Forms.RadioButton();
            this.LabelReadPath = new System.Windows.Forms.Label();
            this.LabelDeliveryPath = new System.Windows.Forms.Label();
            this.LabelPath = new System.Windows.Forms.Label();
            this.ButtonReadPath = new System.Windows.Forms.Button();
            this.LabelDeliveryPathTitle = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.ButtonDeliveryPath = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.LabelDeliveryAction = new System.Windows.Forms.Label();
            this.ButtonSendPath = new System.Windows.Forms.Button();
            this.LabelSendPathTitle = new System.Windows.Forms.Label();
            this.LabelSendPathValue = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.CheckBoxUseSamePathDelivery = new System.Windows.Forms.CheckBox();
            this.treeView = new System.Windows.Forms.TreeView();
            this.PanelSend.SuspendLayout();
            this.PanelRead.SuspendLayout();
            this.PanelDelivery.SuspendLayout();
            this.SuspendLayout();
            // 
            // PanelSend
            // 
            this.PanelSend.Controls.Add(this.RadioButtonSendDelete);
            this.PanelSend.Controls.Add(this.RadioButtonSendFile);
            this.PanelSend.Controls.Add(this.RadioButtonSendNoAction);
            this.PanelSend.Location = new System.Drawing.Point(50, 393);
            this.PanelSend.Name = "PanelSend";
            this.PanelSend.Size = new System.Drawing.Size(189, 107);
            this.PanelSend.TabIndex = 157;
            // 
            // RadioButtonSendDelete
            // 
            this.RadioButtonSendDelete.AutoSize = true;
            this.RadioButtonSendDelete.Location = new System.Drawing.Point(9, 41);
            this.RadioButtonSendDelete.Name = "RadioButtonSendDelete";
            this.RadioButtonSendDelete.Size = new System.Drawing.Size(112, 24);
            this.RadioButtonSendDelete.TabIndex = 59;
            this.RadioButtonSendDelete.Text = "Delete Email";
            this.RadioButtonSendDelete.UseVisualStyleBackColor = true;
            this.RadioButtonSendDelete.CheckedChanged += new System.EventHandler(this.RadioButtonSendDelete_CheckedChanged);
            // 
            // RadioButtonSendFile
            // 
            this.RadioButtonSendFile.AutoSize = true;
            this.RadioButtonSendFile.Location = new System.Drawing.Point(9, 74);
            this.RadioButtonSendFile.Name = "RadioButtonSendFile";
            this.RadioButtonSendFile.Size = new System.Drawing.Size(155, 24);
            this.RadioButtonSendFile.TabIndex = 58;
            this.RadioButtonSendFile.Text = "File Email to Folder";
            this.RadioButtonSendFile.UseVisualStyleBackColor = true;
            this.RadioButtonSendFile.CheckedChanged += new System.EventHandler(this.RadioButtonSendFile_CheckedChanged);
            // 
            // RadioButtonSendNoAction
            // 
            this.RadioButtonSendNoAction.AutoSize = true;
            this.RadioButtonSendNoAction.Checked = true;
            this.RadioButtonSendNoAction.Location = new System.Drawing.Point(9, 8);
            this.RadioButtonSendNoAction.Name = "RadioButtonSendNoAction";
            this.RadioButtonSendNoAction.Size = new System.Drawing.Size(94, 24);
            this.RadioButtonSendNoAction.TabIndex = 57;
            this.RadioButtonSendNoAction.TabStop = true;
            this.RadioButtonSendNoAction.Text = "No Action";
            this.RadioButtonSendNoAction.UseVisualStyleBackColor = true;
            this.RadioButtonSendNoAction.CheckedChanged += new System.EventHandler(this.RadioButtonSendNoAction_CheckedChanged);
            // 
            // PanelRead
            // 
            this.PanelRead.Controls.Add(this.RadioButtonReadJunk);
            this.PanelRead.Controls.Add(this.RadioButtonReadDelete);
            this.PanelRead.Controls.Add(this.RadioButtonReadFile);
            this.PanelRead.Controls.Add(this.RadioButtonReadNoAction);
            this.PanelRead.Location = new System.Drawing.Point(52, 219);
            this.PanelRead.Name = "PanelRead";
            this.PanelRead.Size = new System.Drawing.Size(189, 138);
            this.PanelRead.TabIndex = 156;
            // 
            // RadioButtonReadJunk
            // 
            this.RadioButtonReadJunk.AutoSize = true;
            this.RadioButtonReadJunk.Location = new System.Drawing.Point(9, 104);
            this.RadioButtonReadJunk.Name = "RadioButtonReadJunk";
            this.RadioButtonReadJunk.Size = new System.Drawing.Size(155, 24);
            this.RadioButtonReadJunk.TabIndex = 44;
            this.RadioButtonReadJunk.Text = "Move email to Junk";
            this.RadioButtonReadJunk.UseVisualStyleBackColor = true;
            this.RadioButtonReadJunk.CheckedChanged += new System.EventHandler(this.RadioButtonReadJunk_CheckedChanged);
            // 
            // RadioButtonReadDelete
            // 
            this.RadioButtonReadDelete.AutoSize = true;
            this.RadioButtonReadDelete.Location = new System.Drawing.Point(9, 41);
            this.RadioButtonReadDelete.Name = "RadioButtonReadDelete";
            this.RadioButtonReadDelete.Size = new System.Drawing.Size(112, 24);
            this.RadioButtonReadDelete.TabIndex = 43;
            this.RadioButtonReadDelete.Text = "Delete Email";
            this.RadioButtonReadDelete.UseVisualStyleBackColor = true;
            this.RadioButtonReadDelete.CheckedChanged += new System.EventHandler(this.RadioButtonReadDelete_CheckedChanged);
            // 
            // RadioButtonReadFile
            // 
            this.RadioButtonReadFile.AutoSize = true;
            this.RadioButtonReadFile.Location = new System.Drawing.Point(9, 74);
            this.RadioButtonReadFile.Name = "RadioButtonReadFile";
            this.RadioButtonReadFile.Size = new System.Drawing.Size(155, 24);
            this.RadioButtonReadFile.TabIndex = 42;
            this.RadioButtonReadFile.Text = "File Email to Folder";
            this.RadioButtonReadFile.UseVisualStyleBackColor = true;
            this.RadioButtonReadFile.CheckedChanged += new System.EventHandler(this.RadioButtonReadFile_CheckedChanged);
            // 
            // RadioButtonReadNoAction
            // 
            this.RadioButtonReadNoAction.AutoSize = true;
            this.RadioButtonReadNoAction.Checked = true;
            this.RadioButtonReadNoAction.Location = new System.Drawing.Point(9, 8);
            this.RadioButtonReadNoAction.Name = "RadioButtonReadNoAction";
            this.RadioButtonReadNoAction.Size = new System.Drawing.Size(94, 24);
            this.RadioButtonReadNoAction.TabIndex = 41;
            this.RadioButtonReadNoAction.TabStop = true;
            this.RadioButtonReadNoAction.Text = "No Action";
            this.RadioButtonReadNoAction.UseVisualStyleBackColor = true;
            this.RadioButtonReadNoAction.CheckedChanged += new System.EventHandler(this.RadioButtonReadNoAction_CheckedChanged);
            // 
            // PanelDelivery
            // 
            this.PanelDelivery.Controls.Add(this.RadioButtonDeliveryJunk);
            this.PanelDelivery.Controls.Add(this.RadioButtonDeliveryDelete);
            this.PanelDelivery.Controls.Add(this.RadioButtonDeliveryFile);
            this.PanelDelivery.Controls.Add(this.RadioButtonDeliveryNoAction);
            this.PanelDelivery.Location = new System.Drawing.Point(50, 42);
            this.PanelDelivery.Name = "PanelDelivery";
            this.PanelDelivery.Size = new System.Drawing.Size(189, 137);
            this.PanelDelivery.TabIndex = 155;
            // 
            // RadioButtonDeliveryJunk
            // 
            this.RadioButtonDeliveryJunk.AutoSize = true;
            this.RadioButtonDeliveryJunk.Location = new System.Drawing.Point(9, 100);
            this.RadioButtonDeliveryJunk.Name = "RadioButtonDeliveryJunk";
            this.RadioButtonDeliveryJunk.Size = new System.Drawing.Size(155, 24);
            this.RadioButtonDeliveryJunk.TabIndex = 35;
            this.RadioButtonDeliveryJunk.Text = "Move email to Junk";
            this.RadioButtonDeliveryJunk.UseVisualStyleBackColor = true;
            this.RadioButtonDeliveryJunk.CheckedChanged += new System.EventHandler(this.RadioButtonDeliveryJunk_CheckedChanged);
            // 
            // RadioButtonDeliveryDelete
            // 
            this.RadioButtonDeliveryDelete.AutoSize = true;
            this.RadioButtonDeliveryDelete.Location = new System.Drawing.Point(9, 40);
            this.RadioButtonDeliveryDelete.Name = "RadioButtonDeliveryDelete";
            this.RadioButtonDeliveryDelete.Size = new System.Drawing.Size(112, 24);
            this.RadioButtonDeliveryDelete.TabIndex = 34;
            this.RadioButtonDeliveryDelete.Text = "Delete Email";
            this.RadioButtonDeliveryDelete.UseVisualStyleBackColor = true;
            this.RadioButtonDeliveryDelete.CheckedChanged += new System.EventHandler(this.RadioButtonDeliveryDelete_CheckedChanged);
            // 
            // RadioButtonDeliveryFile
            // 
            this.RadioButtonDeliveryFile.AutoSize = true;
            this.RadioButtonDeliveryFile.Location = new System.Drawing.Point(9, 70);
            this.RadioButtonDeliveryFile.Name = "RadioButtonDeliveryFile";
            this.RadioButtonDeliveryFile.Size = new System.Drawing.Size(167, 24);
            this.RadioButtonDeliveryFile.TabIndex = 33;
            this.RadioButtonDeliveryFile.Text = "Move email to folder";
            this.RadioButtonDeliveryFile.UseVisualStyleBackColor = true;
            this.RadioButtonDeliveryFile.CheckedChanged += new System.EventHandler(this.RadioButtonDeliveryFile_CheckedChanged);
            // 
            // RadioButtonDeliveryNoAction
            // 
            this.RadioButtonDeliveryNoAction.AutoSize = true;
            this.RadioButtonDeliveryNoAction.Checked = true;
            this.RadioButtonDeliveryNoAction.Location = new System.Drawing.Point(9, 10);
            this.RadioButtonDeliveryNoAction.Name = "RadioButtonDeliveryNoAction";
            this.RadioButtonDeliveryNoAction.Size = new System.Drawing.Size(94, 24);
            this.RadioButtonDeliveryNoAction.TabIndex = 32;
            this.RadioButtonDeliveryNoAction.TabStop = true;
            this.RadioButtonDeliveryNoAction.Text = "No Action";
            this.RadioButtonDeliveryNoAction.UseVisualStyleBackColor = true;
            this.RadioButtonDeliveryNoAction.CheckedChanged += new System.EventHandler(this.RadioButtonDeliveryNoAction_CheckedChanged);
            // 
            // LabelReadPath
            // 
            this.LabelReadPath.AutoSize = true;
            this.LabelReadPath.Location = new System.Drawing.Point(327, 295);
            this.LabelReadPath.Name = "LabelReadPath";
            this.LabelReadPath.Size = new System.Drawing.Size(15, 20);
            this.LabelReadPath.TabIndex = 153;
            this.LabelReadPath.Text = "/";
            // 
            // LabelDeliveryPath
            // 
            this.LabelDeliveryPath.AutoSize = true;
            this.LabelDeliveryPath.Location = new System.Drawing.Point(325, 120);
            this.LabelDeliveryPath.Name = "LabelDeliveryPath";
            this.LabelDeliveryPath.Size = new System.Drawing.Size(15, 20);
            this.LabelDeliveryPath.TabIndex = 152;
            this.LabelDeliveryPath.Text = "/";
            this.LabelDeliveryPath.Visible = false;
            // 
            // LabelPath
            // 
            this.LabelPath.AutoSize = true;
            this.LabelPath.Location = new System.Drawing.Point(281, 295);
            this.LabelPath.Name = "LabelPath";
            this.LabelPath.Size = new System.Drawing.Size(40, 20);
            this.LabelPath.TabIndex = 151;
            this.LabelPath.Text = "Path:";
            // 
            // ButtonReadPath
            // 
            this.ButtonReadPath.Location = new System.Drawing.Point(247, 292);
            this.ButtonReadPath.Name = "ButtonReadPath";
            this.ButtonReadPath.Size = new System.Drawing.Size(28, 27);
            this.ButtonReadPath.TabIndex = 150;
            this.ButtonReadPath.Text = "...";
            this.ButtonReadPath.UseVisualStyleBackColor = true;
            this.ButtonReadPath.Click += new System.EventHandler(this.ButtonReadPath_Click);
            // 
            // LabelDeliveryPathTitle
            // 
            this.LabelDeliveryPathTitle.AutoSize = true;
            this.LabelDeliveryPathTitle.Location = new System.Drawing.Point(279, 120);
            this.LabelDeliveryPathTitle.Name = "LabelDeliveryPathTitle";
            this.LabelDeliveryPathTitle.Size = new System.Drawing.Size(40, 20);
            this.LabelDeliveryPathTitle.TabIndex = 149;
            this.LabelDeliveryPathTitle.Text = "Path:";
            this.LabelDeliveryPathTitle.Visible = false;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label15.Location = new System.Drawing.Point(20, 372);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(151, 18);
            this.label15.TabIndex = 143;
            this.label15.Text = "After Send Action";
            // 
            // ButtonDeliveryPath
            // 
            this.ButtonDeliveryPath.Location = new System.Drawing.Point(245, 117);
            this.ButtonDeliveryPath.Name = "ButtonDeliveryPath";
            this.ButtonDeliveryPath.Size = new System.Drawing.Size(28, 27);
            this.ButtonDeliveryPath.TabIndex = 134;
            this.ButtonDeliveryPath.Text = "...";
            this.ButtonDeliveryPath.UseVisualStyleBackColor = true;
            this.ButtonDeliveryPath.Visible = false;
            this.ButtonDeliveryPath.Click += new System.EventHandler(this.ButtonDeliveryPath_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(21, 198);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(152, 18);
            this.label13.TabIndex = 127;
            this.label13.Text = "After Read Action";
            // 
            // LabelDeliveryAction
            // 
            this.LabelDeliveryAction.AutoSize = true;
            this.LabelDeliveryAction.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LabelDeliveryAction.Location = new System.Drawing.Point(21, 21);
            this.LabelDeliveryAction.Name = "LabelDeliveryAction";
            this.LabelDeliveryAction.Size = new System.Drawing.Size(177, 18);
            this.LabelDeliveryAction.TabIndex = 126;
            this.LabelDeliveryAction.Text = "After Delivery Action";
            // 
            // ButtonSendPath
            // 
            this.ButtonSendPath.Location = new System.Drawing.Point(245, 466);
            this.ButtonSendPath.Name = "ButtonSendPath";
            this.ButtonSendPath.Size = new System.Drawing.Size(28, 27);
            this.ButtonSendPath.TabIndex = 158;
            this.ButtonSendPath.Text = "...";
            this.ButtonSendPath.UseVisualStyleBackColor = true;
            this.ButtonSendPath.Click += new System.EventHandler(this.ButtonSendPath_Click);
            // 
            // LabelSendPathTitle
            // 
            this.LabelSendPathTitle.AutoSize = true;
            this.LabelSendPathTitle.Location = new System.Drawing.Point(279, 469);
            this.LabelSendPathTitle.Name = "LabelSendPathTitle";
            this.LabelSendPathTitle.Size = new System.Drawing.Size(40, 20);
            this.LabelSendPathTitle.TabIndex = 159;
            this.LabelSendPathTitle.Text = "Path:";
            // 
            // LabelSendPathValue
            // 
            this.LabelSendPathValue.AutoSize = true;
            this.LabelSendPathValue.Location = new System.Drawing.Point(325, 469);
            this.LabelSendPathValue.Name = "LabelSendPathValue";
            this.LabelSendPathValue.Size = new System.Drawing.Size(15, 20);
            this.LabelSendPathValue.TabIndex = 160;
            this.LabelSendPathValue.Text = "/";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(217, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(354, 20);
            this.label1.TabIndex = 162;
            this.label1.Text = "(What to do after an email is delivered to the Inbox)";
            this.label1.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(219, 198);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(348, 20);
            this.label2.TabIndex = 163;
            this.label2.Text = "(What to do after an email an email has been read)";
            this.label2.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(217, 371);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(284, 20);
            this.label3.TabIndex = 164;
            this.label3.Text = "(What to do after an email has been sent)";
            this.label3.Visible = false;
            // 
            // CheckBoxUseSamePathDelivery
            // 
            this.CheckBoxUseSamePathDelivery.AutoSize = true;
            this.CheckBoxUseSamePathDelivery.Checked = true;
            this.CheckBoxUseSamePathDelivery.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CheckBoxUseSamePathDelivery.Location = new System.Drawing.Point(61, 506);
            this.CheckBoxUseSamePathDelivery.Name = "CheckBoxUseSamePathDelivery";
            this.CheckBoxUseSamePathDelivery.Size = new System.Drawing.Size(368, 24);
            this.CheckBoxUseSamePathDelivery.TabIndex = 165;
            this.CheckBoxUseSamePathDelivery.Text = "Use a similar path in both the Inbox and Sent Items.";
            this.CheckBoxUseSamePathDelivery.UseVisualStyleBackColor = true;
            this.CheckBoxUseSamePathDelivery.Visible = false;
            this.CheckBoxUseSamePathDelivery.CheckedChanged += new System.EventHandler(this.CheckBoxUseSamePathDelivery_CheckedChanged);
            // 
            // treeView
            // 
            this.treeView.Location = new System.Drawing.Point(590, 82);
            this.treeView.Name = "treeView";
            treeNode1.Name = "Node0";
            treeNode1.Text = "InBox";
            this.treeView.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1});
            this.treeView.Size = new System.Drawing.Size(486, 448);
            this.treeView.TabIndex = 166;
            this.treeView.Visible = false;
            this.treeView.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // ContactInTouchSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.treeView);
            this.Controls.Add(this.CheckBoxUseSamePathDelivery);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.LabelSendPathValue);
            this.Controls.Add(this.LabelSendPathTitle);
            this.Controls.Add(this.ButtonSendPath);
            this.Controls.Add(this.PanelSend);
            this.Controls.Add(this.PanelRead);
            this.Controls.Add(this.PanelDelivery);
            this.Controls.Add(this.LabelReadPath);
            this.Controls.Add(this.LabelDeliveryPath);
            this.Controls.Add(this.LabelPath);
            this.Controls.Add(this.ButtonReadPath);
            this.Controls.Add(this.LabelDeliveryPathTitle);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.ButtonDeliveryPath);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.LabelDeliveryAction);
            this.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "ContactInTouchSettings";
            this.Size = new System.Drawing.Size(1105, 653);
            this.FormRegionShowing += new System.EventHandler(this.ContactInTouchSettings_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.ContactInTouchSettings_FormRegionClosed);
            this.PanelSend.ResumeLayout(false);
            this.PanelSend.PerformLayout();
            this.PanelRead.ResumeLayout(false);
            this.PanelRead.PerformLayout();
            this.PanelDelivery.ResumeLayout(false);
            this.PanelDelivery.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel PanelSend;
        private System.Windows.Forms.RadioButton RadioButtonSendDelete;
        private System.Windows.Forms.RadioButton RadioButtonSendFile;
        private System.Windows.Forms.RadioButton RadioButtonSendNoAction;
        private System.Windows.Forms.Panel PanelRead;
        private System.Windows.Forms.RadioButton RadioButtonReadDelete;
        private System.Windows.Forms.RadioButton RadioButtonReadFile;
        private System.Windows.Forms.RadioButton RadioButtonReadNoAction;
        private System.Windows.Forms.Panel PanelDelivery;
        private System.Windows.Forms.RadioButton RadioButtonDeliveryDelete;
        private System.Windows.Forms.RadioButton RadioButtonDeliveryFile;
        private System.Windows.Forms.RadioButton RadioButtonDeliveryNoAction;
        private System.Windows.Forms.Label LabelReadPath;
        private System.Windows.Forms.Label LabelDeliveryPath;
        private System.Windows.Forms.Label LabelPath;
        private System.Windows.Forms.Button ButtonReadPath;
        private System.Windows.Forms.Label LabelDeliveryPathTitle;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button ButtonDeliveryPath;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label LabelDeliveryAction;
        private System.Windows.Forms.Button ButtonSendPath;
        private System.Windows.Forms.Label LabelSendPathTitle;
        private System.Windows.Forms.Label LabelSendPathValue;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox CheckBoxUseSamePathDelivery;
        private System.Windows.Forms.RadioButton RadioButtonDeliveryJunk;
        private System.Windows.Forms.RadioButton RadioButtonReadJunk;
        private System.Windows.Forms.TreeView treeView;

        public partial class ContactInTouchSettingsFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public ContactInTouchSettingsFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                ContactInTouchSettings.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.ContactInTouchSettingsFactory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                ContactInTouchSettings form = new ContactInTouchSettings(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                throw new System.NotSupportedException();
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.WindowsForms;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal ContactInTouchSettings ContactInTouchSettings
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(ContactInTouchSettings))
                        return (ContactInTouchSettings)item;
                }
                return null;
            }
        }
    }
}
