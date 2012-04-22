using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace ExampleBase
{
    partial class FormError : Form
    {
        public FormError(string title, string message, Exception exception)
        {
            InitializeComponent();
            if (null == message)
                message = FormOptions.LCID == 1033 ? "An error is occured." : "Ein Fehler ist aufgetreten.";

            this.Text = title;
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

        public static void Show(Control parentDialog, Exception exception)
        {
            FormError form = new FormError(FormOptions.LCID == 1033 ? "Error." : "Fehler.", null, exception);
            form.ShowDialog(parentDialog);
        }

        public static void Show(Control parentDialog, string title, string message, Exception exception)
        {
            FormError form = new FormError(title, message, exception);

            form.ShowDialog(parentDialog);
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
