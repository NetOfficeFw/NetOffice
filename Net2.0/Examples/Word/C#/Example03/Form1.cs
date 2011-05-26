using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums; 

namespace Example03
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;
             
            // add a new document
            Word.Document newDocument = wordApplication.Documents.Add();

            // create a new listtemplate
            Word.ListTemplate template = newDocument.ListTemplates.Add(true, "NetOfficeListTemplate");

            //get the predefined listlevels (9)
            Word.ListLevels levels = template.ListLevels;

            // customize the first level of the list
            levels[1].NumberFormat = "%1.";

            // tab is used to change the level
            levels[1].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            levels[1].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
            levels[1].NumberPosition = 0;
            levels[1].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;    
            levels[1].TextPosition = wordApplication.CentimetersToPoints(0.63F);
            levels[1].TabPosition = wordApplication.CentimetersToPoints(0.63F);
            levels[1].ResetOnHigher = 0;
            levels[1].StartAt = 1;
            levels[1].LinkedStyle = "";
            levels[1].Font.Bold = 1;

            // customize the second level of the list
            levels[2].NumberFormat = "%1.%2.";

            // tab is used to change the level
            levels[2].TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            levels[2].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;

            // we want the numbers to appear under the first letter of the higher level
            levels[2].NumberPosition = wordApplication.CentimetersToPoints(0.63F);
            levels[2].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;

            // and the text should indent a tab more on the right
            levels[2].TextPosition = wordApplication.CentimetersToPoints(1.4F);
            levels[2].TabPosition = wordApplication.CentimetersToPoints(1.4F);
            levels[2].ResetOnHigher = 0;
            levels[2].StartAt = 1;
            levels[2].LinkedStyle = "";
            levels[2].Font.Italic = 1;

            // apply the defined listtemplate to the selection
            wordApplication.Selection.Range.ListFormat.ApplyListTemplate(template, false, 
                    WdListApplyTo.wdListApplyToWholeList, WdDefaultListBehavior.wdWord9ListBehavior);

            //create a list
            wordApplication.Selection.TypeText("Welcoming");
            wordApplication.Selection.TypeParagraph();

            wordApplication.Selection.TypeText("Introduction");
            wordApplication.Selection.TypeParagraph();

            wordApplication.Selection.TypeText("Presentation");
            wordApplication.Selection.TypeParagraph();

            // execute the indent so the second level gets activated
            wordApplication.Selection.Range.ListFormat.ListIndent();

            wordApplication.Selection.TypeText("Top 1");
            wordApplication.Selection.TypeParagraph();

            wordApplication.Selection.TypeText("Top 2");
            wordApplication.Selection.TypeParagraph();

            wordApplication.Selection.TypeText("Top 3");
            wordApplication.Selection.TypeParagraph();

            // execute the outdent so the first level gets reactivated
            wordApplication.Selection.Range.ListFormat.ListOutdent();
            wordApplication.Selection.TypeText("Questions & Answers");

            //save the document
            string fileExtension = GetDefaultExtension(wordApplication);
            object documentFile = string.Format("{0}\\Example03{1}", Application.StartupPath, fileExtension);
            newDocument.SaveAs(documentFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Document saved.", documentFile.ToString());
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
            double version = Convert.ToDouble(application.Version);
            if (version >= 120.00)
                return ".docx";
            else
                return ".doc";
        }

        #endregion
    }
}
