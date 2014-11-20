using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.Custom
{
    public partial class InviteControl : UserControl
    {
        public InviteControl()
        {
            InitializeComponent();
            controlBackColorAnimator1.Start(false);
        }

        private void InviteControl_Resize(object sender, EventArgs e)
        {
            panelBottom.Left = (this.Width / 2) - (panelBottom.Width / 2);
            panelBottom.Top = this.Height - panelBottom.Height;
        }

        private void linkLabelMail_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(linkLabelMail.Text);
            }
            catch
            {
                ;
            }
        }

        private void linkLabelFolder_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                if (!Directory.Exists(Translation.ToolLanguages.DirectoryPath))
                    Directory.CreateDirectory(Translation.ToolLanguages.DirectoryPath);
                System.Diagnostics.Process.Start(Translation.ToolLanguages.DirectoryPath);
            }
            catch
            {
                ;    
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Visible = false;
        }
    }
}
