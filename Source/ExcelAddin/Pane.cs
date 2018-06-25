using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.ExcelApi;
using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace ExcelAddin
{
    public partial class Pane : UserControl, NetOffice.ExcelApi.Tools.ITaskPane
    {
        public Pane()
        {
            InitializeComponent();
        }

        public void OnConnection(NetOffice.ExcelApi.Application application, _CustomTaskPane parentPane, object[] customArguments)
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
