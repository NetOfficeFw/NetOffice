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
    /// Example 1 - Create a presentation
    /// </summary>
    internal class Example01 : IExample 
    {
        #region IExample

        public void RunExample()
        {
            // start powerpoint
            PowerPoint.Application powerApplication = new PowerPoint.Application();

            // create a utils instance, not need for but helpful to keep the lines of code low
            PowerPoint.Tools.CommonUtils utils = new PowerPoint.Tools.CommonUtils(powerApplication);

            // add a new presentation with one new slide
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            presentation.Slides.Add(1, PpSlideLayout.ppLayoutClipArtAndVerticalText);
            
            // save the document 
            string documentFile = utils.File.Combine(HostApplication.RootDirectory, "Example01", PowerPoint.Tools.DocumentFormat.Normal);
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
            get { return HostApplication.LCID == 1033 ? "Example01" : "Beispiel01"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create a presentation with 1 empty slide" : "Eine neue Präsentation erstellen"; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion
    }
}
