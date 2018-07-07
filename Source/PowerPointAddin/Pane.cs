using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using NetOffice.PowerPointApi;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace PowerPointAddin
{
    public partial class Pane : UserControl, NetOffice.OfficeApi.Tools.ITaskPane
    {
        public void OnConnection(ICOMObject application, _CustomTaskPane parentPane, object[] customArguments)
        {

        }

        public void OnDisconnection()
        {

        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {

        }

        public void OnVisibleStateChanged(bool visible)
        {

        }
    }
}
