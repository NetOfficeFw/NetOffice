using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using Core = LateBindingApi.Core;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelTests
{
    /// <summary>
    /// shapes
    /// </summary>
    public class Test04 :ITestPackage
    {
        #region ITestPackage Member

        public bool DoTest(string logFilePath)
        {
            Core.DebugConsole.FileName = System.IO.Path.Combine(logFilePath, "ExcelTests.Test04.log");
            Core.DebugConsole.AppendTimeInfoEnabled = true;
            Core.DebugConsole.Mode = LateBindingApi.Core.ConsoleMode.LogFile;
           
            Excel.Application application = null;
            try
            {
                //  Initialize NetOffice
                LateBindingApi.Core.Factory.Initialize();

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

                return true;
            }
            catch (Exception exception)
            {
                string message = exception.Message;
                Console.WriteLine("An error occured{1}{1}{0}", message, Environment.NewLine);
                return false;
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
