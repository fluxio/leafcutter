namespace GoogleDriveGH.Components
{
    partial class AuthForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.authLink = new System.Windows.Forms.LinkLabel();
            this.tokenTextBox = new System.Windows.Forms.TextBox();
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(393, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Please click the link below to authenticate with Google.";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(418, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "Copy the resulting code into the text box below and hit OK.";
            // 
            // authLink
            // 
            this.authLink.AutoSize = true;
            this.authLink.Location = new System.Drawing.Point(167, 85);
            this.authLink.Name = "authLink";
            this.authLink.Size = new System.Drawing.Size(135, 20);
            this.authLink.TabIndex = 2;
            this.authLink.TabStop = true;
            this.authLink.Text = "Authenticate Now";
            this.authLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.authLink_LinkClicked);
            // 
            // tokenTextBox
            // 
            this.tokenTextBox.Location = new System.Drawing.Point(29, 118);
            this.tokenTextBox.Name = "tokenTextBox";
            this.tokenTextBox.Size = new System.Drawing.Size(414, 26);
            this.tokenTextBox.TabIndex = 3;
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(288, 176);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 27);
            this.cancelButton.TabIndex = 4;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // okButton
            // 
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.okButton.Location = new System.Drawing.Point(379, 176);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(75, 27);
            this.okButton.TabIndex = 5;
            this.okButton.Text = "OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.okButton_Click);
            // 
            // AuthForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(476, 215);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.tokenTextBox);
            this.Controls.Add(this.authLink);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "AuthForm";
            this.Text = "Google Authorization Form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.LinkLabel authLink;
        private System.Windows.Forms.TextBox tokenTextBox;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button okButton;
    }
}