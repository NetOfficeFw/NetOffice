using System;
using System.Windows.Forms;
using ExampleBase;
using NetOffice;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace PowerPointExamplesCS4
{
    /// <summary>
    /// Example 6 - Using events
    /// </summary>
    internal partial class Example06 : UserControl, IExample
    {
        #region Fields/Delegates

        private delegate void UpdateEventTextDelegate(string Message);
        private UpdateEventTextDelegate _updateDelegate;

        #endregion

        #region Ctor

        public Example06()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        #endregion

        #region IExample

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example06"; }
        }

        public string Description
        {
            get { return "Using Events"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Methods

        private void UpdateTextbox(string message)
        {
            textBoxEvents.AppendText(message + "\r\n");
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

        private void powerApplication_PresentationCloseEvent(NetOffice.PowerPointApi.Presentation Pres)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event PresentationClose called." });
            Pres.Dispose();
        }

        private void powerApplication_AfterNewPresentationEvent(NetOffice.PowerPointApi.Presentation Pres)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event AfterNewPresentation called." });
            Pres.Dispose();
        }

        #endregion
    }
}
