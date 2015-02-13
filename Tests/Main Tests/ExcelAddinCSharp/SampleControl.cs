using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelAddinCSharp
{
    public partial class SampleControl : UserControl, NetOffice.ExcelApi.Tools.ITaskPane
    {
        public SampleControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Hello");
        }

        void NetOffice.ExcelApi.Tools.ITaskPane.OnConnection(NetOffice.ExcelApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            TestAddin addin = customArguments[0] as TestAddin;
            addin.TaskPaneOkay = true;
        }

        void NetOffice.ExcelApi.Tools.ITaskPane.OnDisconnection()
        {
             
        }

        void NetOffice.ExcelApi.Tools.ITaskPane.OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position)
        {
           
        }

        void NetOffice.ExcelApi.Tools.ITaskPane.OnVisibleStateChanged(bool visible)
        {

        }
    }
}
