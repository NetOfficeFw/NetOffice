using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Text;
using System.Globalization;
using ExampleBase;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.VBIDEApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace ExcelExamplesCS4
{
    class Example04 : IExample
    {
        IHost _hostApplication;

        #region IExample Member

        public void RunExample()
        {          
            // Initialize NetOffice
            LateBindingApi.Core.Factory.Initialize();

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
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

            // save the book 
            string fileExtension = GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example04{1}", _hostApplication.RootDirectory, fileExtension);
            workBook.SaveAs(workbookFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, NetOffice.ExcelApi.Enums.XlSaveAsAccessMode.xlExclusive);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            // show dialog for the user(you!)
            _hostApplication.ShowFinishDialog(null, workbookFile);
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Shapes, WordArts, Pictures, 3D-Effects" : "Shapes, WordArts, Pictures, 3D-Effects"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".xlsx";
            else
                return ".xls";
        }

        #endregion
    }
}
