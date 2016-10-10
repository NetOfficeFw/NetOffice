using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.Tools;

namespace ClientAddin
{
    public static class MessageDialog
    {
        private static bool _enabled = false;

        public static void ShowRegister(Type type, RegisterCall registerCall, string from)
        {
            string text = String.Format("Type:{0}{2}RegisterCall:{0}{2}", type, registerCall, Environment.NewLine);

            if (_enabled)
                MessageBox.Show(text, "Register " + from);
            else
                Console.WriteLine("Register " + from + " :" + text);
        }

        public static void ShowRegister(InstallScope scope, string from)
        {
            string text = String.Format("Scope:{0}", scope);

            if (_enabled)
                MessageBox.Show(text, "Register " + from);
            else
                Console.WriteLine("Register " + from + " :" + text);
        }

        public static void ShowRegister(Type type, RegisterCall registerCall, InstallScope scope, OfficeRegisterKeyState keyState, string from)
        {
            string text = String.Format(
                "Register Type:{1}{0}Assembly Name:{2}{0}Assembly Version:{3}{0}RegisterCall:{4}{0}Scope:{5}{0}KeyState:{6}{0}",
                Environment.NewLine, type, type.Assembly.GetName().Name, type.Assembly.GetName().Version, registerCall, scope, keyState);

            if (_enabled)
                MessageBox.Show(text, "Register " + from);
            else
                Console.WriteLine("Register " + from + " :" + text);
        }

        public static void ShowUnRegister(Type type, RegisterCall registerCall, string from)
        {
            string text = String.Format("Type:{0}{2}RegisterCall:{0}{2}", type, registerCall, Environment.NewLine);

            if (_enabled)
                MessageBox.Show(text, "Unregister " + from);
            else
                Console.WriteLine("Unregister " + from + " :" + text);
        }

        public static void ShowUnRegister(InstallScope scope, string from)
        {
            string text = String.Format("Scope:{0}", scope);

            if (_enabled)
                MessageBox.Show(text, "Unregister " + from);
            else
                Console.WriteLine("Unregister " + from + " :" + text);
        }

        public static void ShowUnRegister(Type type, RegisterCall registerCall, InstallScope scope, OfficeUnRegisterKeyState keyState, string from)
        {
            string text = String.Format(
                 "Register Type:{1}{0}Assembly Name:{2}{0}Assembly Version:{3}{0}RegisterCall:{4}{0}Scope:{5}{0}KeyState:{6}{0}",
                 Environment.NewLine, type, type.Assembly.GetName().Name, type.Assembly.GetName().Version, registerCall, scope, keyState);

            if (_enabled)
                MessageBox.Show(text, "Unregister " + from);
            else
                Console.WriteLine("Unregister " + from + " :" + text);
        }
        
        public static void ShowRegisterError(RegisterErrorMethodKind methodKind, Exception exception, string from)
        {
            string text = String.Format("Method:{0}{2}{2}{1}", methodKind, exception, Environment.NewLine);
            MessageBox.Show(text, "Reg/Unreg Error " + from);
        }
    }
}
