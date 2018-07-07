using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

namespace InnerAddin
{
    /*
     *
     * TODO: provide optional interface and base class
    */
    [ComVisible(true)]
    [ProgId("InnerAddin.InnerAddinPane")]
    [Guid("E702FB12-92C4-4DBA-8848-45134BFD3448")]
    public partial class InnerAddinPane : UserControl
    {
        public InnerAddinPane()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            MessageBox.Show(
                "This is button 1." + Environment.NewLine +
                "AppDomain HashCode:" + AppDomain.CurrentDomain.GetHashCode().ToString()
                );
        }
    }
}
