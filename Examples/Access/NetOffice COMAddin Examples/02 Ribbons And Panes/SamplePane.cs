using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace Access02AddinCS4
{
    public partial class SamplePane : UserControl, NetOffice.AccessApi.Tools.ITaskPane // Not necessary to implement ITaskPane but its helpful
    {
        #region Ctor

        public SamplePane()
        {           
            InitializeComponent();
        }

        #endregion

        #region Properties

        private PerformanceCounter Counter { get; set; }

        #endregion

        #region ITaskpane

        public void OnConnection(NetOffice.AccessApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {

        }

        public void OnDisconnection()
        {
            UsageTimer.Enabled = false;
            if (null != Counter)
            {
                Counter.Dispose();
                Counter = null;
            }
        }

        public void OnVisibleStateChanged(bool visible)
        {
            // Create the performance counter is expensive in performance
            // To avoid slow down the Access startup sequence - we create them on demand when user want show the pane
            // (Real world code want doing that async)
            if (visible && null == Counter)
            {
                Counter = new PerformanceCounter("Process", "% Processor Time", "MSACCESS");
                UsageTimer.Enabled = true;
            }
            else if (visible)
                UsageTimer.Enabled = true;
            else if (!visible)
                UsageTimer.Enabled = false;
        }

        public void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position)
        {
            
        }

        #endregion

        #region UI Trigger

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UsageLabel.Location = new Point(
                                    (Width / 2 - UsageLabel.Width / 2),
                                    (Height / 2 - UsageLabel.Height / 2));
        }

        private void UsageTimer_Tick(object sender, EventArgs e)
        {
            if (null != Counter)
            {
                float value = Counter.NextValue();
              
                int barValue = Convert.ToInt32(value);
                if (barValue < 0)
                    barValue = 0;
                if (barValue > 100)
                    barValue = 100;
                UsageLabel.Text = String.Format("{0} %", barValue);
                UsageBar.Value = barValue;
            }
            else
            {
                UsageLabel.Text = String.Empty;
            }
        }

        #endregion
    }
}