using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOfficeTools.ExtendedOutlookCS4
{
    public partial class SamplePane : UserControl, NetOffice.OutlookApi.Tools.ITaskPane
    {
        #region Ctor

        public SamplePane()
        {
            InitializeComponent();
        }

        #endregion

        #region Properties

        private DateTime StartTime { get; set; }

        #endregion

        #region ITaskpane

        public void OnConnection(NetOffice.OutlookApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            StartTime = DateTime.Now;
            buttonEnabled_Click(buttonEnabled, new EventArgs());
        }

        public void OnDisconnection()
        {

        }

        #endregion

        #region UI Trigger

        private void buttonEnabled_Click(object sender, EventArgs e)
        {
            if (timerRunningTime.Enabled)
            {
                timerRunningTime.Enabled = false;
                buttonEnabled.Text = "Enable";
                buttonEnabled.ImageKey = "alarmclock_run.png";
                labelTime.ForeColor = Color.FromKnownColor(KnownColor.ControlText);
            }
            else
            {
                timerRunningTime.Enabled = true;
                buttonEnabled.Text = "Disable";
                buttonEnabled.ImageKey = "alarmclock_stop.png";
                labelTime.ForeColor = Color.White;
            }
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            StartTime = DateTime.Now;
            labelTime.Text = "00:00:00";
        }

        private void timerRunningTime_Tick(object sender, EventArgs e)
        {
            TimeSpan ts = DateTime.Now - StartTime;
            labelTime.Text = String.Format("{0:00}:{1:00}:{2:00}", ts.Hours, ts.Minutes, ts.Seconds);

        }

        #endregion
    }
}
