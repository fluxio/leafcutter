using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GoogleDriveGH.Components
{
    public partial class AuthForm : Form
    {
        public string token;
        private string launchURL;
        public AuthForm()
        {
            launchURL = "";
            token = "";
            InitializeComponent();
        }

        public DialogResult ShowURL(string URL, IWin32Window owner)
        {
            launchURL = URL;
            return this.ShowDialog(owner);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void okButton_Click(object sender, EventArgs e)
        {
            token = tokenTextBox.Text;
            this.Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void authLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(launchURL);

        }
    }
}
