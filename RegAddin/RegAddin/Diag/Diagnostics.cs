using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RegAddin.Diag
{
    internal class Diagnostics
    {
        internal void Show(IEnumerable<string> args)
        {
            StringBuilder messageBuilder = new StringBuilder();
            messageBuilder.AppendLine("Command Arguments:");
            if (null != args)
            { 
                foreach (string item in args)
                    messageBuilder.AppendLine(" - " + item);
            }
            else
                messageBuilder.AppendLine("<Null Arguments Reference>");

            messageBuilder.AppendLine(Environment.NewLine + "Application Path:");
            messageBuilder.AppendLine(GetCodebase());
            messageBuilder.AppendLine(Environment.NewLine +  "Version:" + About.AssemblyVersion);
            messageBuilder.AppendLine(Environment.NewLine);
            messageBuilder.AppendLine("Press Ctrl+C to copy this message into Clipboard.");

            MessageBox.Show(messageBuilder.ToString(), About.AssemblyTitle + " Diagnostics", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GetCodebase()
        {
            string removePreQuote = "file:///";
            string codebase = typeof(Diagnostics).Assembly.CodeBase;
            if (codebase.StartsWith(removePreQuote, StringComparison.InvariantCultureIgnoreCase))
                codebase = codebase.Substring(removePreQuote.Length);
            codebase = codebase.Replace("/", "\\");
            return codebase;
        }
    }
}
