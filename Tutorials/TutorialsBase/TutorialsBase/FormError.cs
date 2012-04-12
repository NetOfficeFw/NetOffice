using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace TutorialsBase
{
    partial class FormError : Form
    {
        public FormError(string title, string message, Exception exception)
        {
            InitializeComponent();
            this.Text = title;
            if (null == message)
                message = "An Error occured";
            labelErrorMessage.Text = message;
            DisplayException(exception);
        }

        private void DisplayException(Exception exception)
        {
            listViewTrace.Items.Clear();

            if (null == exception)
                return;

            int i = 1;
            while (exception != null)
            {
                ListViewItem viewItem = listViewTrace.Items.Add(i.ToString());
                viewItem.SubItems.Add(exception.Message);
                viewItem.SubItems.Add(exception.GetType().Name.ToString());
                if (null != exception.TargetSite)
                    viewItem.SubItems.Add(exception.TargetSite.ToString());
                else
                    viewItem.SubItems.Add("");
                exception = exception.InnerException;
                i++;
            }
        }

        public static void Show(Control parentDialog, string title, string message, Exception exception)
        {
            if (title == null)
                title = FormOptions.LCID == 1031 ? "An error is occured." : "Ein Fehler ist aufgetreten.";

            FormError form = new FormError(title, message, exception);
            form.ShowDialog(parentDialog);
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
