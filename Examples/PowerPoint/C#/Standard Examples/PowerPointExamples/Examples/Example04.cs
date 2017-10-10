using System;
using System.Windows.Forms;
using System.Globalization;
using ExampleBase;
using NetOffice;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;
using NetOffice.PowerPointApi.Tools.Contribution;

namespace PowerPointExamplesCS4
{
    /// <summary>
    /// Example 4 - Create blend animation
    /// </summary>
    internal class Example04 : IExample
    {
        public void RunExample()
        {
            // start powerpoint 
            PowerPoint.Application powerApplication = new PowerPoint.Application();
         
            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(powerApplication);

            // add a new presentation with two new slides
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.Slide slide1 = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
            PowerPoint.Slide slide2 = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add shapes
            slide1.Shapes.AddShape(MsoAutoShapeType.msoShape4pointStar, 100, 100, 200, 200);
            slide2.Shapes.AddShape(MsoAutoShapeType.msoShapeDoubleWave, 200, 200, 200, 200);

            // change blend animation
            slide1.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverDown;
            slide1.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast;

            slide2.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverLeftDown;
            slide2.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast;

            // save the document 
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example04", DocumentFormat.Normal); 
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
            get { return "Example04"; }
        }

        public string Description
        {
            get { return "Create blend animation"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
