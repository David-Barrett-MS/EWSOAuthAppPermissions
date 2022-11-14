
namespace EWSOAuthAppPermissions
{
    partial class Form1
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxClientSecret = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxTenantId = new System.Windows.Forms.TextBox();
            this.textBoxAppId = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.textBoxResults = new System.Windows.Forms.TextBox();
            this.buttonClose = new System.Windows.Forms.Button();
            this.buttonFindFolders = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.checkBoxPOXBasicAuth = new System.Windows.Forms.CheckBox();
            this.textBoxAutoDiscoverPW = new System.Windows.Forms.TextBox();
            this.checkBoxGetPublicFolders = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxMailboxSMTPAddress = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBoxClientSecret);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBoxTenantId);
            this.groupBox1.Controls.Add(this.textBoxAppId);
            this.groupBox1.Location = new System.Drawing.Point(18, 18);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox1.Size = new System.Drawing.Size(490, 152);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Application Information";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 114);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 20);
            this.label3.TabIndex = 5;
            this.label3.Text = "Client Secret:";
            // 
            // textBoxClientSecret
            // 
            this.textBoxClientSecret.Location = new System.Drawing.Point(129, 109);
            this.textBoxClientSecret.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxClientSecret.Name = "textBoxClientSecret";
            this.textBoxClientSecret.Size = new System.Drawing.Size(350, 26);
            this.textBoxClientSecret.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 74);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "Tenant Id:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 34);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(109, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "Application Id:";
            // 
            // textBoxTenantId
            // 
            this.textBoxTenantId.Location = new System.Drawing.Point(129, 69);
            this.textBoxTenantId.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxTenantId.Name = "textBoxTenantId";
            this.textBoxTenantId.Size = new System.Drawing.Size(350, 26);
            this.textBoxTenantId.TabIndex = 1;
            // 
            // textBoxAppId
            // 
            this.textBoxAppId.Location = new System.Drawing.Point(129, 29);
            this.textBoxAppId.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxAppId.Name = "textBoxAppId";
            this.textBoxAppId.Size = new System.Drawing.Size(350, 26);
            this.textBoxAppId.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBoxResults);
            this.groupBox2.Location = new System.Drawing.Point(18, 331);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox2.Size = new System.Drawing.Size(490, 192);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Results";
            // 
            // textBoxResults
            // 
            this.textBoxResults.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxResults.Location = new System.Drawing.Point(4, 24);
            this.textBoxResults.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxResults.Multiline = true;
            this.textBoxResults.Name = "textBoxResults";
            this.textBoxResults.ReadOnly = true;
            this.textBoxResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxResults.Size = new System.Drawing.Size(482, 163);
            this.textBoxResults.TabIndex = 0;
            // 
            // buttonClose
            // 
            this.buttonClose.Location = new System.Drawing.Point(396, 532);
            this.buttonClose.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonClose.Name = "buttonClose";
            this.buttonClose.Size = new System.Drawing.Size(112, 35);
            this.buttonClose.TabIndex = 7;
            this.buttonClose.Text = "Close";
            this.buttonClose.UseVisualStyleBackColor = true;
            this.buttonClose.Click += new System.EventHandler(this.buttonClose_Click);
            // 
            // buttonFindFolders
            // 
            this.buttonFindFolders.Location = new System.Drawing.Point(18, 534);
            this.buttonFindFolders.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.buttonFindFolders.Name = "buttonFindFolders";
            this.buttonFindFolders.Size = new System.Drawing.Size(112, 35);
            this.buttonFindFolders.TabIndex = 6;
            this.buttonFindFolders.Text = "Find Folders";
            this.buttonFindFolders.UseVisualStyleBackColor = true;
            this.buttonFindFolders.Click += new System.EventHandler(this.buttonFindFolders_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBoxPOXBasicAuth);
            this.groupBox3.Controls.Add(this.textBoxAutoDiscoverPW);
            this.groupBox3.Controls.Add(this.checkBoxGetPublicFolders);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Controls.Add(this.textBoxMailboxSMTPAddress);
            this.groupBox3.Location = new System.Drawing.Point(18, 180);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.groupBox3.Size = new System.Drawing.Size(490, 142);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Mailbox";
            // 
            // checkBoxPOXBasicAuth
            // 
            this.checkBoxPOXBasicAuth.AutoSize = true;
            this.checkBoxPOXBasicAuth.Location = new System.Drawing.Point(76, 102);
            this.checkBoxPOXBasicAuth.Name = "checkBoxPOXBasicAuth";
            this.checkBoxPOXBasicAuth.Size = new System.Drawing.Size(201, 24);
            this.checkBoxPOXBasicAuth.TabIndex = 10;
            this.checkBoxPOXBasicAuth.Text = "Use basic auth for POX";
            this.toolTip1.SetToolTip(this.checkBoxPOXBasicAuth, "If selected, POX is used for AutoDiscover to determine X-PublicFolderMailbox (oth" +
        "erwise basic auth is used)");
            this.checkBoxPOXBasicAuth.UseVisualStyleBackColor = true;
            // 
            // textBoxAutoDiscoverPW
            // 
            this.textBoxAutoDiscoverPW.Location = new System.Drawing.Point(288, 97);
            this.textBoxAutoDiscoverPW.Name = "textBoxAutoDiscoverPW";
            this.textBoxAutoDiscoverPW.Size = new System.Drawing.Size(192, 26);
            this.textBoxAutoDiscoverPW.TabIndex = 9;
            this.toolTip1.SetToolTip(this.textBoxAutoDiscoverPW, "Public folder access requires information that can only be\r\nobtained via AutoDisc" +
        "over using basic auth.  Enter the\r\npassword for the mailbox here to ensure publi" +
        "c folder\r\nheaders are set correctly.");
            this.textBoxAutoDiscoverPW.UseSystemPasswordChar = true;
            // 
            // checkBoxGetPublicFolders
            // 
            this.checkBoxGetPublicFolders.AutoSize = true;
            this.checkBoxGetPublicFolders.Location = new System.Drawing.Point(76, 69);
            this.checkBoxGetPublicFolders.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.checkBoxGetPublicFolders.Name = "checkBoxGetPublicFolders";
            this.checkBoxGetPublicFolders.Size = new System.Drawing.Size(191, 24);
            this.checkBoxGetPublicFolders.TabIndex = 8;
            this.checkBoxGetPublicFolders.Text = "Retrieve public folders";
            this.toolTip1.SetToolTip(this.checkBoxGetPublicFolders, "If checked, public folders will be accessed instead of\r\nthe user\'s primary mailbo" +
        "x.");
            this.checkBoxGetPublicFolders.UseVisualStyleBackColor = true;
            this.checkBoxGetPublicFolders.CheckedChanged += new System.EventHandler(this.checkBoxGetPublicFolders_CheckedChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 34);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(117, 20);
            this.label4.TabIndex = 7;
            this.label4.Text = "SMTP address:";
            // 
            // textBoxMailboxSMTPAddress
            // 
            this.textBoxMailboxSMTPAddress.Location = new System.Drawing.Point(129, 29);
            this.textBoxMailboxSMTPAddress.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.textBoxMailboxSMTPAddress.Name = "textBoxMailboxSMTPAddress";
            this.textBoxMailboxSMTPAddress.Size = new System.Drawing.Size(350, 26);
            this.textBoxMailboxSMTPAddress.TabIndex = 6;
            this.toolTip1.SetToolTip(this.textBoxMailboxSMTPAddress, "The SMTP address of the mailbox being impersonated/accessed.");
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 583);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.buttonClose);
            this.Controls.Add(this.buttonFindFolders);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "EWS OAuth App Permissions Sample (using MSAL)";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxTenantId;
        private System.Windows.Forms.TextBox textBoxAppId;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox textBoxResults;
        private System.Windows.Forms.Button buttonClose;
        private System.Windows.Forms.Button buttonFindFolders;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxClientSecret;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxMailboxSMTPAddress;
        private System.Windows.Forms.CheckBox checkBoxGetPublicFolders;
        private System.Windows.Forms.TextBox textBoxAutoDiscoverPW;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox checkBoxPOXBasicAuth;
    }
}

