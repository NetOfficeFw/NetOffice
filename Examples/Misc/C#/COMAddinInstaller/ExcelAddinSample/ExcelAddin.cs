using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NetOffice.ExcelApi.Tools;

namespace ExcelAddinSample
{
    [ProgId("NetOffice.ExcelAddinInstaller")]
    [Guid("121A6EA7-AD5F-471B-A750-F4D6280472C0")]
    public class ExcelAddin : COMAddin
    {

        public ExcelAddin()
        {
            this.OnStartupComplete += ExcelAddin_OnStartupComplete;
        }

        private void ExcelAddin_OnStartupComplete(ref Array custom)
        {
            MessageBox.Show($"NetOffice Add-in running in Microsoft Excel {this.Application.Version}", "NetOffice Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
