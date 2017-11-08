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
    /// <summary>
    /// A custom control to invite people for translations
    /// </summary>
    public partial class InviteControl : UserControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public InviteControl()
        {
            InitializeComponent();
            controlBackColorAnimator1.Start(false);
        }

        #endregion

        #region Trigger

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
                
            }
            catch
            {
                ;    
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Visible = false;
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(null, exception);
            }
        }

        #endregion
    }
}
