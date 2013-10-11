using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOfficeTools.ExtendedPPointCS4
{
    public partial class SamplePane : UserControl, NetOffice.PowerPointApi.Tools.ITaskPane
    {
        public SamplePane()
        {
            InitializeComponent();
        }

        void NetOffice.PowerPointApi.Tools.ITaskPane.OnConnection(NetOffice.PowerPointApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {

        }

        void NetOffice.PowerPointApi.Tools.ITaskPane.OnDisconnection()
        {

        }
    }
}
