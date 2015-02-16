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
    /// Example 4 - Create blend animation
    /// </summary>
    internal class Example04 : IExample
    {
        #region IExample

        public void RunExample()
        {
            // start powerpoint 
            PowerPoint.Application powerApplication = new PowerPoint.Application();

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
            string fileExtension = GetDefaultExtension(powerApplication);
            string documentFile = string.Format("{0}\\Example04{1}", HostApplication.RootDirectory, fileExtension);
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
            get { return HostApplication.LCID == 1033 ? "Example04" : "Beispiel04"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create blend animation" : "Eine Blend Animation erstellen"; }
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
