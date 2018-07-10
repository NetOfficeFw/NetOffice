using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using NetOffice.Tools.Isolation;

namespace InnerUpdate
{
    public class Connect : IManagedInnerUpdateHandler
    {
        public IShimUpdateHost Parent { get; private set; }

        public object Application { get; private set; }

        public string Custom { get; private set; }

        public void SetParent([In, MarshalAs(UnmanagedType.IUnknown)] IShimUpdateHost shim)
        {
            if (null == shim)
                throw new ArgumentNullException("shim");
            Parent = shim;
        }

        public void SetCustomData([In, MarshalAs(UnmanagedType.BStr)] string custom)
        {
            Custom = custom;
        }

        public void SetApplication([In, MarshalAs(UnmanagedType.IDispatch)] object application)
        {
            Application = application;

            //string applicationName = TypeDescriptor.GetClassName(application);
            //MessageBox.Show(String.Format("Execute Applications is {0}", applicationName), "IManagedInnerUpdateHandler::Connect");
        }

        public void CanExecute([In, MarshalAs(UnmanagedType.Bool), Out] ref bool canExecute)
        {
            canExecute = true;
        }

        public void Execute()
        {
            Trace.WriteLine("IManagedInnerUpdateHandler::Execute::Begin");
            Thread.Sleep(1000);
            Trace.WriteLine("IManagedInnerUpdateHandler::Execute::End");
        }

        public void Close()
        {
            bool free = false;
            if (null != Application)
            {
                Marshal.ReleaseComObject(Application);
                Application = null;
                free = true;
            }
            if (null != Parent)
            {
                Parent.SetCustomData(Custom + " " + "Greetings too from Update Handler.");
                Marshal.ReleaseComObject(Parent);
                Parent = null;
                free = true;
            }

            if (free)
            {
                GC.Collect();
                GC.Collect();
            }

            //MessageBox.Show("Thanks for close notify me.", "IManagedInnerUpdateHandler::Connect");
        }
    }
}
