using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RegAddin.Alert
{
    internal static class Window
    {
        internal static void ShowSucceedMessage()
        {
            string message = null;
            switch (SingletonSettings.Mode)
            {                 
                case SingletonSettings.ApplicationMode.Register:
                    message = "Register Operation Successfully.";
                    break;
                case SingletonSettings.ApplicationMode.Unregister:
                    message = "Unregister Operation Successfully.";
                    break;
                case SingletonSettings.ApplicationMode.RegFile:
                    message = "Register Export Successfully.";
                    break;
            }
            if(null != message)
                Show(message);
        }

        internal static void Show(string text)
        {
            MessageBox.Show(text, About.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal static void Show(string text, params object[] args)
        {
            MessageBox.Show(String.Format(text, args), About.AssemblyTitle);
        } 

        internal static void ShowError(string text)
        {
            MessageBox.Show(text, About.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        internal static void ShowError(string text, params object[] args)
        {
            MessageBox.Show(String.Format(text, args), About.AssemblyTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
