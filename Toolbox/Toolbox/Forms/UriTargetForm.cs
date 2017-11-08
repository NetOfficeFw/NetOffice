using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Forms
{
    public partial class UriTargetForm : Form
    {
        public UriTargetForm()
        {
            InitializeComponent();
        }

        public string SelectedUri
        {
            get
            {
                return OsdnRadioButton.Checked == true ? OsdnRadioButton.Tag as string : GithubRadioButton.Tag as string;
            }
        }

        private Label LastClickedLinkLabel { get; set; }

        public static string ShowForm(IWin32Window owner)
        {
            UriTargetForm form = new UriTargetForm();
            if (form.ShowDialog(owner) == DialogResult.OK)
            {
                string result = form.SelectedUri;
                form.Dispose();
                return result;
            }
            else
                return null;
        }

        private void OsdnLinkLabel_Click(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null == mouseArgs)
                    return;

                if (mouseArgs.Button == MouseButtons.Right)
                {
                    LastClickedLinkLabel = sender as Label;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void UriTargetForm_Click(object sender, EventArgs args)
        {
            try
            {
                MouseEventArgs mouseArgs = args as MouseEventArgs;
                if (null == mouseArgs)
                    return;

                if (mouseArgs.Button == MouseButtons.Right)
                {
                    LastClickedLinkLabel = sender as Label;
                    LinkContextMenu.Show(sender as Control, 0, 0);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(this, exception, ErrorCategory.NonCritical);
            }
        }

        private void LinkContextMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (null != LastClickedLinkLabel)
            {
                Clipboard.SetText(LastClickedLinkLabel.Tag as string);
            }
        }
    }
}