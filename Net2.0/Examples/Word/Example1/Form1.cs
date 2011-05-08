using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Word = LateBindingApi.WordApi;
using LateBindingApi.WordApi.Enums; 

namespace Example1
{
    public partial class Form1 : Form
    {
        Word.Application _wordApplication;

        public Form1()
        {
            InitializeComponent();

            /*
             * Initialize Api
             */
            LateBindingApi.Core.Factory.Initialize();
        }
  
        private void button1_Click(object sender, EventArgs e)
        {

            // start word and turn off msg boxes
            _wordApplication = new Word.Application();
            _wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
 
            // add a new document
            Word.Document newDocument = _wordApplication.Documents.Add();

            // save the document 
            object missingValue = Missing.Value;
            string fileExtension = GetDefaultExtension(_wordApplication);
            object documentFile = string.Format("{0}\\Example01{1}", Environment.CurrentDirectory, fileExtension);
            newDocument.SaveAs(ref documentFile, ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                                ref missingValue, ref missingValue, ref missingValue, ref missingValue, ref missingValue,
                                ref missingValue, ref missingValue, ref missingValue, ref missingValue, ref missingValue, ref missingValue); 

            // close word and dispose reference
            _wordApplication.Quit();
            _wordApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Document saved.", (string)documentFile);
            fDialog.ShowDialog(this);
        }
        
        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".doc" or ".docx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Word.Application application)
        {
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".docx";
            else
                return ".doc";
        }
        
        #endregion
    }
}
