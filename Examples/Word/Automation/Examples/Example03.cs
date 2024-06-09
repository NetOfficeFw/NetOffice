﻿using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools.Contribution;

namespace WordExamplesCS4
{
    /// <summary>
    /// Example 3 - Using templates
    /// </summary>
    internal class Example03 : IExample
    {
        public void RunExample()
        {
            // start word and turn off msg boxes
            Word.Application wordApplication = new Word.Application();
            wordApplication.DisplayAlerts = WdAlertLevel.wdAlertsNone;

            // create a utils instance, not need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(wordApplication);

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

            // save the document
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example03", DocumentFormat.Normal);
            newDocument.SaveAs(documentFile);

            // close word and dispose reference
            wordApplication.Quit();
            wordApplication.Dispose();

            // show end dialog
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example03"; }
        }

        public string Description
        {
            get { return "Using templates"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }
        
        internal IHost HostApplication { get; private set; }
    }
}
