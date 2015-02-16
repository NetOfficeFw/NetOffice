using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using ExampleBase;

using NetOffice;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace PowerPointExamplesCS4
{
    /// <summary>
    /// Example 2 - Create shapes
    /// </summary>
    internal class Example02 : IExample
    {
        #region IExample

        public void RunExample()
        {
            // start powerpoint 
            PowerPoint.Application powerApplication = new PowerPoint.Application();

            // add a new presentation with one new slide
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add a label
            PowerPoint.Shape label = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 600, 20);
            label.TextFrame.TextRange.Text = "This slide and created Shapes are created by NetOffice example.";

            // add a line
            slide.Shapes.AddLine(10, 80, 700, 80);

            // add a wordart
            slide.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect9, "This a WordArt", "Arial", 20,
                                           MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 150);

            // add a star
            slide.Shapes.AddShape(MsoAutoShapeType.msoShape24pointStar, 200, 200, 250, 250);

            // save the document 
            string fileExtension = GetDefaultExtension(powerApplication);
            string documentFile = string.Format("{0}\\Example02{1}", HostApplication.RootDirectory, fileExtension);
            presentation.SaveAs(documentFile);

            // close power point and dispose reference
            powerApplication.Quit();
            powerApplication.Dispose();

            // show dialog for the user(you!)
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example02" : "Beispiel02"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create some kind of shapes" : "Verschiede Shapes erstellen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".ppt" or ".pptx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(PowerPoint.Application application)
        {
            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".pptx";
            else
                return ".ppt";
        }

        #endregion
    }
}
