using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Tools.Contribution;

namespace PowerPointExamplesCS4
{
    /// <summary>
    /// Example 2 - Create shapes
    /// </summary>
    internal class Example02 : IExample
    {
        public void RunExample()
        {
            // start powerpoint 
            PowerPoint.Application powerApplication = new PowerPoint.Application();

            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(powerApplication);

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
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example02", DocumentFormat.Normal); 
            presentation.SaveAs(documentFile);

            // close power point and dispose reference
            powerApplication.Quit();
            powerApplication.Dispose();

            // show end dialog
            HostApplication.ShowFinishDialog(null, documentFile);
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example02"; }
        }

        public string Description
        {
            get { return "Create some kind of shapes"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
