using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOfficeTools.ExtendedAccessCS4
{
    public partial class SamplePane : UserControl, NetOffice.AccessApi.Tools.ITaskPane
    {
        public SamplePane()
        {
            InitializeComponent();
        }

        void NetOffice.AccessApi.Tools.ITaskPane.OnConnection(NetOffice.AccessApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {

        }

        void NetOffice.AccessApi.Tools.ITaskPane.OnDisconnection()
        {
             
        }
    }
}
