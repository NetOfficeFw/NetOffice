using System;
using System.Windows.Forms;
using System.Globalization;
using ExampleBase;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.OfficeApi.Enums;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 4 - Shapes, WordArts, Pictures, 3D-Effects
    /// </summary>
    internal class Example04 : IExample
    {
        #region IExample

        public void RunExample()
        {          
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;

            // create a utils instance, not need for but helpful to keep the lines of code low
            Excel.Tools.CommonUtils utils = new Excel.Tools.CommonUtils(excelApplication);

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            workSheet.Cells[1, 1].Value = "These sample shapes was dynamicly created by code.";

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
            string workbookFile = utils.File.Combine(HostApplication.RootDirectory, "Example04", Excel.Tools.DocumentFormat.Normal);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            // show dialog for the user(you!)
            HostApplication.ShowFinishDialog(null, workbookFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Shapes, WordArts, Pictures, 3D-Effects" : "Shapes, WordArts, Pictures, 3D-Effects"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Current Example Host
        /// </summary>
        internal IHost HostApplication { get; private set; }

        #endregion
    }
}
