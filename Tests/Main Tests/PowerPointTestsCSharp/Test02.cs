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
    public class Test02 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test02"; }
        }

        public string Description
        {
            get { return "Create WordArts."; }
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
