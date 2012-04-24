using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelTestsCSharp
{
    public class Test04 :ITestPackage
    {
        #region ITestPackage Member
         
        public string Name
        {
            get { return "Test04"; }
        }

        public string Description
        {
            get { return "Using shapes"; }
        }

        public string OfficeProduct
        {
            get { return "Excel"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Excel.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                // start excel and turn off msg boxes
                application = new Excel.Application();
                application.DisplayAlerts = false;

                // add a new workbook
                Excel.Workbook workBook = application.Workbooks.Add();
                Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

                workSheet.Cells[1, 1].Value = "these sample shapes was dynamicly created by code.";

                // create a star
                Excel.Shape starShape = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShape32pointStar, 10, 50, 200, 20);

                // create a simple textbox
                Excel.Shape textBox = workSheet.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 150, 200, 50);
                textBox.TextFrame.Characters().Text = "text";
                textBox.TextFrame.Characters().Font.Size = 14;

                // create a wordart
                Excel.Shape textEffect = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect14, "WordArt", "Arial", 12,
                                                                                    MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 250);

                // create text effect
                Excel.Shape textDiagram = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14,
                                                                                    MsoTriState.msoFalse, MsoTriState.msoFalse, 10, 350);

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null,  "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }
        }

        #endregion
    }
}
