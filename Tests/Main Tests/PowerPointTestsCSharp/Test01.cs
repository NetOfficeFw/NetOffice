using System;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;

namespace PowerPointTestsCSharp
{
    public class Test01 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test01"; }
        }

        public string Description
        {
            get { return "Create a presentation."; }
        }

        public string OfficeProduct
        {
            get { return "PowerPoint"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            PowerPoint.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                application = new PowerPoint.Application();

                // add a new presentation with one new slide
                PowerPoint.Presentation presentation = application.Presentations.Add(MsoTriState.msoTrue);
                presentation.Slides.Add(1, PpSlideLayout.ppLayoutClipArtAndVerticalText);  

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
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
