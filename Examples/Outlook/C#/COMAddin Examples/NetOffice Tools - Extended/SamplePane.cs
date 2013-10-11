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
        public SamplePane()
        {
            InitializeComponent();
        }

        void NetOffice.OutlookApi.Tools.ITaskPane.OnConnection(NetOffice.OutlookApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {

        }

        void NetOffice.OutlookApi.Tools.ITaskPane.OnDisconnection()
        {

        }
    }
}
