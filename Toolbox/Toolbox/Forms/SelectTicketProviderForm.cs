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
    public partial class SelectTicketProviderForm : Form
    {
        public SelectTicketProviderForm()
        {
            InitializeComponent();
        }

        public static void ShowForm(IWin32Window owner)
        {
            SelectTicketProviderForm form = new SelectTicketProviderForm();
            form.ShowDialog(owner);
        }

        private void ProceedButton_Click(object sender, EventArgs e)
        {
            TryStartUri(GetUri());
            Close();
        }

        private void AbortButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private string GetUri()
        {
            if (OsdnBox.Checked)
                return OsdnBox.Tag as string;
            else
                return GithubBox.Tag as string;
        }

        private static void TryStartUri(string uri)
        {
            try
            {
                System.Diagnostics.Process.Start(uri);
            }
            catch
            {
                ;
            }
        }
    }
}