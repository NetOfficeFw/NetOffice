using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;

namespace Example06
{
    public partial class Form1 : Form
    {
        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start word and turn off msg boxes
            PowerPoint.Application powerApplication = new PowerPoint.Application();
            powerApplication.DisplayAlerts = PpAlertLevel.ppAlertsNone;

            /*
            we register some events. note: the event trigger was called from power point, means an other Thread
            remove the Quit() call below and check out more events if you want
            */

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
    }
}
