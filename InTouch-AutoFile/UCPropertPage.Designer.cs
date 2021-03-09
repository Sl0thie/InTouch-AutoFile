
namespace InTouch_AutoFile
{
    sealed partial class UCPropertPage
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.CheckBoxTaskInbox = new System.Windows.Forms.CheckBox();
            this.CheckBoxTaskSent = new System.Windows.Forms.CheckBox();
            this.CheckBoxTaskDuplicates = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CheckBoxShowTasksButton = new System.Windows.Forms.CheckBox();
            this.CheckBoxTaskRouting = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CheckBoxTaskInbox
            // 
            this.CheckBoxTaskInbox.AutoSize = true;
            this.CheckBoxTaskInbox.Location = new System.Drawing.Point(26, 140);
            this.CheckBoxTaskInbox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CheckBoxTaskInbox.Name = "CheckBoxTaskInbox";
            this.CheckBoxTaskInbox.Size = new System.Drawing.Size(117, 21);
            this.CheckBoxTaskInbox.TabIndex = 0;
            this.CheckBoxTaskInbox.Text = "File Inbox Items";
            this.CheckBoxTaskInbox.UseVisualStyleBackColor = true;
            this.CheckBoxTaskInbox.CheckedChanged += new System.EventHandler(this.CheckBoxTaskInbox_CheckedChanged);
            // 
            // CheckBoxTaskSent
            // 
            this.CheckBoxTaskSent.AutoSize = true;
            this.CheckBoxTaskSent.Location = new System.Drawing.Point(26, 169);
            this.CheckBoxTaskSent.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CheckBoxTaskSent.Name = "CheckBoxTaskSent";
            this.CheckBoxTaskSent.Size = new System.Drawing.Size(110, 21);
            this.CheckBoxTaskSent.TabIndex = 1;
            this.CheckBoxTaskSent.Text = "File Sent Items";
            this.CheckBoxTaskSent.UseVisualStyleBackColor = true;
            this.CheckBoxTaskSent.CheckedChanged += new System.EventHandler(this.CheckBoxTaskSent_CheckedChanged);
            // 
            // CheckBoxTaskDuplicates
            // 
            this.CheckBoxTaskDuplicates.AutoSize = true;
            this.CheckBoxTaskDuplicates.Location = new System.Drawing.Point(26, 198);
            this.CheckBoxTaskDuplicates.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CheckBoxTaskDuplicates.Name = "CheckBoxTaskDuplicates";
            this.CheckBoxTaskDuplicates.Size = new System.Drawing.Size(211, 21);
            this.CheckBoxTaskDuplicates.TabIndex = 2;
            this.CheckBoxTaskDuplicates.Text = "Find Dulplicate email addresses";
            this.CheckBoxTaskDuplicates.UseVisualStyleBackColor = true;
            this.CheckBoxTaskDuplicates.CheckedChanged += new System.EventHandler(this.CheckBoxTaskDuplicates_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(15, 107);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 20);
            this.label1.TabIndex = 3;
            this.label1.Text = "Automated Tasks";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 11);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 20);
            this.label2.TabIndex = 4;
            this.label2.Text = "Options";
            // 
            // CheckBoxShowTasksButton
            // 
            this.CheckBoxShowTasksButton.AutoSize = true;
            this.CheckBoxShowTasksButton.Location = new System.Drawing.Point(26, 35);
            this.CheckBoxShowTasksButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CheckBoxShowTasksButton.Name = "CheckBoxShowTasksButton";
            this.CheckBoxShowTasksButton.Size = new System.Drawing.Size(134, 21);
            this.CheckBoxShowTasksButton.TabIndex = 5;
            this.CheckBoxShowTasksButton.Text = "Show Tasks Button";
            this.CheckBoxShowTasksButton.UseVisualStyleBackColor = true;
            this.CheckBoxShowTasksButton.CheckedChanged += new System.EventHandler(this.CheckBoxShowTasksButton_CheckedChanged);
            // 
            // CheckBoxTaskRouting
            // 
            this.CheckBoxTaskRouting.AutoSize = true;
            this.CheckBoxTaskRouting.Location = new System.Drawing.Point(26, 227);
            this.CheckBoxTaskRouting.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.CheckBoxTaskRouting.Name = "CheckBoxTaskRouting";
            this.CheckBoxTaskRouting.Size = new System.Drawing.Size(145, 21);
            this.CheckBoxTaskRouting.TabIndex = 6;
            this.CheckBoxTaskRouting.Text = "Email Routing Check";
            this.CheckBoxTaskRouting.UseVisualStyleBackColor = true;
            this.CheckBoxTaskRouting.CheckedChanged += new System.EventHandler(this.CheckBoxTaskRouting_CheckedChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(177, 221);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(29, 30);
            this.button1.TabIndex = 7;
            this.button1.Text = "...";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // UCPropertPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.button1);
            this.Controls.Add(this.CheckBoxTaskRouting);
            this.Controls.Add(this.CheckBoxShowTasksButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CheckBoxTaskDuplicates);
            this.Controls.Add(this.CheckBoxTaskSent);
            this.Controls.Add(this.CheckBoxTaskInbox);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "UCPropertPage";
            this.Size = new System.Drawing.Size(352, 289);
            this.Load += new System.EventHandler(this.UCPropertPage_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox CheckBoxTaskInbox;
        private System.Windows.Forms.CheckBox CheckBoxTaskSent;
        private System.Windows.Forms.CheckBox CheckBoxTaskDuplicates;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox CheckBoxShowTasksButton;
        private System.Windows.Forms.CheckBox CheckBoxTaskRouting;
        private System.Windows.Forms.Button button1;
    }
}
