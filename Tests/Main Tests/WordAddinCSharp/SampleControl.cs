using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordAddinCSharp
{
    public partial class SampleControl : UserControl, NetOffice.WordApi.Tools.ITaskPane
    {
        public SampleControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Hello");
        }

        void NetOffice.WordApi.Tools.ITaskPane.OnConnection(NetOffice.WordApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            TestAddin addin = customArguments[0] as TestAddin;
            addin.TaskPaneOkay = true;
        }

        void NetOffice.WordApi.Tools.ITaskPane.OnDisconnection()
        {

        }

        void NetOffice.WordApi.Tools.ITaskPane.OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position)
        {

        }

        void NetOffice.WordApi.Tools.ITaskPane.OnVisibleStateChanged(bool visible)
        {

        }
    }
}
