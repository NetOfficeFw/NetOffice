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
    /// Example 1 - Create a presentation
    /// </summary>
    internal class Example01 : IExample 
    {
        public void RunExample()
        {
            // start powerpoint
            PowerPoint.Application powerApplication = new PowerPoint.Application();

            // create a utils instance, no need for but helpful to keep the lines of code low
            CommonUtils utils = new CommonUtils(powerApplication);

            // add a new presentation with one new slide
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutClipArtAndVerticalText);
            
            // save the document 
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", DocumentFormat.Normal);
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
            get { return "Example01"; }
        }

        public string Description
        {
            get { return "Create a presentation with 1 empty slide"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        internal IHost HostApplication { get; private set; }
    }
}
