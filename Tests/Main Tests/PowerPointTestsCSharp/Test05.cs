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
    public class Test05 : ITestPackage
    {
        bool _presentationClose;
        bool _afterNewPresentation;

        #region TestPackage Member

        public string Name
        {
            get { return "Test05."; }
        }

        public string Description
        {
            get { return "Using events."; }
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
                application.Visible = MsoTriState.msoTrue;

                // PowerPoint 2000 doesnt support DisplayAlerts, we check at runtime its available and set
                if (application.EntityIsAvailable("DisplayAlerts"))
                    application.DisplayAlerts = PpAlertLevel.ppAlertsNone;

                application.PresentationCloseEvent += new NetOffice.PowerPointApi.Application_PresentationCloseEventHandler(powerApplication_PresentationCloseEvent);
                application.AfterNewPresentationEvent += new NetOffice.PowerPointApi.Application_AfterNewPresentationEventHandler(powerApplication_AfterNewPresentationEvent);

                // add a new presentation with one new slide
                PowerPoint.Presentation presentation = application.Presentations.Add(MsoTriState.msoTrue);
                PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

                System.Threading.Thread.Sleep(2000);

                // close the document
                presentation.Close();

                if(_afterNewPresentation && _presentationClose)
                    return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
                else
                    return new TestResult(false, DateTime.Now.Subtract(startTime), String.Format("AfterNewPresentation:{0} , PresentationClose:{1}", _afterNewPresentation, _presentationClose), null, "");
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
         
        void powerApplication_PresentationCloseEvent(NetOffice.PowerPointApi.Presentation Pres)
        {
           _presentationClose = true;
            Pres.Dispose();
        }

        void powerApplication_AfterNewPresentationEvent(NetOffice.PowerPointApi.Presentation Pres)
        {
            _afterNewPresentation = true;
            Pres.Dispose();
        }

        #endregion
    }
}
