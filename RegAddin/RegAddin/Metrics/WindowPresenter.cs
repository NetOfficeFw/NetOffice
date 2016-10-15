using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RegAddin.Metrics
{
    internal class WindowPresenter
    {
        internal void Show(Dictionary<string, bool> result)
        {
            StringBuilder message = new StringBuilder();
            message.AppendLine("The following rules failed to validate:");
            foreach (var item in result)
            {
                if (item.Value == false)
                    message.AppendLine(item.Key);
            }

            MessageBox.Show(message.ToString(), About.AssemblyTitle + " Metrics");
        }
    }
}
