using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    partial class FinishDialog : Form
    {
        string _folder;

        public FinishDialog(string folder)
        {
            _folder = folder;
            InitializeComponent();
            if (ProjectWizardControl.CurrentLanguageID == 1031)
            {
                this.Text = "Vorgang abgeschlossen";
                labelCaption.Text = "Das Projekt wurde erfolgreich erstellt.";
                buttonOpen.Text = "Ordner öffnen";
                buttonOK.Text = "Schliessen";
            }
            else
            {
                this.Text = "Succeed";
                labelCaption.Text = "The project is done.";
                buttonOpen.Text = "Open Folder";
                buttonOK.Text = "Close";
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(_folder);
            }
            catch 
            {
                ;
            }
        }
    }
}
