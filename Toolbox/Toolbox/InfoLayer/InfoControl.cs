using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using  System.Reflection;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox
{
    public partial class InfoControl : UserControl
    {

        public InfoControl(string text)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;
            richTextBox.Text = text;            
        }


        public InfoControl(string text, bool isRessourceAddress)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;

            if (isRessourceAddress)
            {
                richTextBox.LoadFile(ReadStream(text), RichTextBoxStreamType.RichText);
            }
            else
            { 
                richTextBox.Text = text;
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private static Stream ReadStream(string resId)
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            System.IO.Stream ressourceStream = ass.GetManifestResourceStream(assemblyName + "." + resId);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            return ressourceStream;
        }

       
        private void richTextBox_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
            catch
            {
                ;
            }
        }

    }

}
