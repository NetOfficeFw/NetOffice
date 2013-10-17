using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordTestsCSharp
{
    public class Test03 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test03"; }
        }

        public string Description
        {
            get { return "Using List templates."; }
        }

        public string OfficeProduct
        {
            get { return "Word"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Word.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new Word.Application();
                application.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                // add a new document
                Word.Document newDocument = application.Documents.Add();

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

                levels[1].TextPosition = application.CentimetersToPoints(0.63F);
                levels[1].TabPosition = application.CentimetersToPoints(0.63F);
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
                levels[2].NumberPosition = application.CentimetersToPoints(0.63F);
                levels[2].Alignment = WdListLevelAlignment.wdListLevelAlignLeft;

                // and the text should indent a tab more on the right
                levels[2].TextPosition = application.CentimetersToPoints(1.4F);
                levels[2].TabPosition = application.CentimetersToPoints(1.4F);
                levels[2].ResetOnHigher = 0;
                levels[2].StartAt = 1;
                levels[2].LinkedStyle = "";
                levels[2].Font.Italic = 1;

                // apply the defined listtemplate to the selection
                application.Selection.Range.ListFormat.ApplyListTemplate(template, false, WdListApplyTo.wdListApplyToWholeList, WdDefaultListBehavior.wdWord9ListBehavior);

                //create a list
                application.Selection.TypeText("Welcoming");
                application.Selection.TypeParagraph();

                application.Selection.TypeText("Introduction");
                application.Selection.TypeParagraph();

                application.Selection.TypeText("Presentation");
                application.Selection.TypeParagraph();

                // execute the indent so the second level gets activated
                application.Selection.Range.ListFormat.ListIndent();

                application.Selection.TypeText("Top 1");
                application.Selection.TypeParagraph();

                application.Selection.TypeText("Top 2");
                application.Selection.TypeParagraph();

                application.Selection.TypeText("Top 3");
                application.Selection.TypeParagraph();

                // execute the outdent so the first level gets reactivated
                application.Selection.Range.ListFormat.ListOutdent();
                application.Selection.TypeText("Questions & Answers");

                newDocument.Close(false);

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit(WdSaveOptions.wdDoNotSaveChanges);
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
