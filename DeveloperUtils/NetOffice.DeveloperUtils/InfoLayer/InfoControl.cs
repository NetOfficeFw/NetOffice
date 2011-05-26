using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using  System.Reflection;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperUtils
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
                richTextBox.Text = ReadString(text);
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


        private static string ReadString(string resId)
        {
            Assembly ass = Assembly.GetExecutingAssembly();

            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            System.IO.Stream ressourceStream = ass.GetManifestResourceStream(assemblyName +"."+ resId);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource File."));

            string text = textStreamReader.ReadToEnd();
            text = text.Replace("\r\n", " ");

            if (null != textStreamReader)
                textStreamReader.Close();
            if (null != ressourceStream)
                ressourceStream.Close();

            return text;
        }

    }

}
