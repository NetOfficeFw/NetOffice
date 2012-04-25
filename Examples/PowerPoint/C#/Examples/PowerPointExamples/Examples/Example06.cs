using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using ExampleBase;

using NetOffice;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace PowerPointExamplesCS4
{
    partial class Example06 : UserControl, IExample
    {
        IHost _hostApplication;

        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Example06()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        #region IExample Member

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            _hostApplication = hostApplication;
        }

        public string Caption
        {
            get { return _hostApplication.LCID == 1033 ? "Example06" : "Beispiel06"; }
        }

        public string Description
        {
            get { return _hostApplication.LCID == 1033 ? "Using Events" : "Verwenden von Ereignissen"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {
            // start powerpoint and turn off msg boxes
            PowerPoint.Application powerApplication = new PowerPoint.Application();
	        powerApplication.Visible = MsoTriState.msoTrue;

            // PowerPoint 2000 doesnt support DisplayAlerts, we check at runtime its available and set
            if(powerApplication.EntityIsAvailable("DisplayAlerts"))
                powerApplication.DisplayAlerts = PpAlertLevel.ppAlertsNone;

            // we register some events. note: the event trigger was called from power point, means an other Thread
            powerApplication.PresentationCloseEvent += new NetOffice.PowerPointApi.Application_PresentationCloseEventHandler(powerApplication_PresentationCloseEvent);
            powerApplication.AfterNewPresentationEvent += new NetOffice.PowerPointApi.Application_AfterNewPresentationEventHandler(powerApplication_AfterNewPresentationEvent);
             
            // add a new presentation with one new slide
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // close the document
            presentation.Close();

            // close power point and dispose reference
            powerApplication.Quit();
            powerApplication.Dispose();
        }

        #endregion

        #region PowerPoint Trigger

        void powerApplication_PresentationCloseEvent(NetOffice.PowerPointApi.Presentation Pres)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event PresentationClose called." });
            Pres.Dispose();
        }

        void powerApplication_AfterNewPresentationEvent(NetOffice.PowerPointApi.Presentation Pres)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event AfterNewPresentation called." });
            Pres.Dispose();
        }

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
        }

        #endregion
    }
}
