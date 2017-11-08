using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using  System.Reflection;
using System.IO;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Controls.InfoLayer
{
    /// <summary>
    /// Control to display rich text
    /// </summary>
    [RessourceTable("Controls.InfoLayer.Strings.txt")]
    public partial class InfoControl : UserControl
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public InfoControl()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="text"></param>
        public InfoControl(string text)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;
            richTextBoxHelpContent.Text = text;            
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="text">rtf text</param>
        /// <param name="isRessourceAddress">indicates first argument is a resource address instead of rtf text</param>
        public InfoControl(string text, bool isRessourceAddress)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;

            if (isRessourceAddress)
            {
                richTextBoxHelpContent.LoadFile(ReadStream(text), RichTextBoxStreamType.RichText);
            }
            else
            {
                richTextBoxHelpContent.Text = text;
            }
        }

        /// <summary>
        /// Stream with rtf data inside
        /// </summary>
        /// <param name="rtfStream">rtf data stream</param>
        public InfoControl(Stream rtfStream)
        {
            InitializeComponent();
            this.Dock = DockStyle.Fill;
            richTextBoxHelpContent.LoadFile(rtfStream, RichTextBoxStreamType.RichText);
        }

        #endregion

        #region Methods

        private static Stream ReadStream(string resId)
        {
            Assembly ass = Assembly.GetExecutingAssembly();
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            System.IO.Stream ressourceStream = ass.GetManifestResourceStream(assemblyName + "." + resId);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));
            return ressourceStream;
        }
            
        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Hide();
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(null, exception);
            }
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

        #endregion
    }
}
