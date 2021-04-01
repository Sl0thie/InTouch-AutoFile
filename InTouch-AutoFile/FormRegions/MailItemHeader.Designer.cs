
namespace InTouch_AutoFile
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class MailItemHeader : Microsoft.Office.Tools.Outlook.FormRegionBase
    {
        public MailItemHeader(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
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
            manifest.FormRegionName = "Message Header";
            manifest.Icons.Page = global::InTouch_AutoFile.Properties.Resources.options;
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
            this.RichText = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // RichText
            // 
            this.RichText.Location = new System.Drawing.Point(0, 0);
            this.RichText.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.RichText.Name = "RichText";
            this.RichText.ReadOnly = true;
            this.RichText.Size = new System.Drawing.Size(505, 367);
            this.RichText.TabIndex = 0;
            this.RichText.Text = "";
            // 
            // MailItemHeader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.RichText);
            this.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "MailItemHeader";
            this.Size = new System.Drawing.Size(681, 708);
            this.FormRegionShowing += new System.EventHandler(this.MailItemHeader_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.MailItemHeader_FormRegionClosed);
            this.Resize += new System.EventHandler(this.MailItemHeader_Resize);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox RichText;

        public partial class MailItemHeaderFactory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public MailItemHeaderFactory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                MailItemHeader.InitializeManifest(this._Manifest, Globals.Factory);
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.MailItemHeaderFactory_FormRegionInitializing);
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
                MailItemHeader form = new MailItemHeader(formRegion);
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
        internal MailItemHeader MailItemHeader
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(MailItemHeader))
                        return (MailItemHeader)item;
                }
                return null;
            }
        }
    }
}
