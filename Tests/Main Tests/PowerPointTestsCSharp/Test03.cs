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
    public class Test03 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test03"; }
        }

        public string Description
        {
            get { return "Create blend animation."; }
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

                // add a new presentation with two new slides
                PowerPoint.Presentation presentation = application.Presentations.Add(MsoTriState.msoTrue);
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
