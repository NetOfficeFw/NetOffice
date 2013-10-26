using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordTestsCSharp
{
    public class Test04 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test04"; }
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

                // create simple a csv-file as datasource
                string fileName = string.Format("{0}\\DataSource.csv", Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
                
                // if file exists then delete
                if (File.Exists(fileName))
                    File.Delete(fileName);

                File.AppendAllText(fileName, string.Format("{0},{1}{2}", "ProjectName", "ProjectLink", Environment.NewLine));
                File.AppendAllText(fileName, string.Format("{0},{1}{2}", "NetOffice", "http://netoffice.codeplex.com/", Environment.NewLine));

                // add a new document
                Word.Document newDocument = application.Documents.Add();

                // define the document as mailmerge
                newDocument.MailMerge.MainDocumentType = WdMailMergeMainDocType.wdFormLetters;

                // open the datasource
                newDocument.MailMerge.OpenDataSource(fileName);

                // insert some text and the mailmergefields defined in the datasource
                application.Selection.TypeText("This test is brought to you by ");
                newDocument.MailMerge.Fields.Add(application.Selection.Range, "ProjectName");

                application.Selection.TypeText(" for more information and examples visit ");
                newDocument.MailMerge.Fields.Add(application.Selection.Range, "ProjectLink ");

                application.Selection.TypeText(" or click ");

                object adress = newDocument.MailMerge.DataSource.DataFields[2].Value;
                object screenTip = "click Tooltip";
                object displayText = "here";
                newDocument.Hyperlinks.Add(application.Selection.Range, adress, Type.Missing, screenTip, displayText, Type.Missing);
                 
                // show the contents of the fields
                int wdToggle = 9999998;
                newDocument.MailMerge.ViewMailMergeFieldCodes = wdToggle;

                //do not show the fieldcodes
                application.ActiveWindow.View.ShowFieldCodes = false;

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
