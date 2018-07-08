using System;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using NetOffice;
using NetOffice.Tools;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.PowerPointApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using VBIDE = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;
using System.Reflection;
using System.Diagnostics;

namespace InnerAddin
{
	[COMAddin("InnerAddin", "InnerAddin Description", LoadBehavior.LoadAtStartup), ProgId("InnerAddin.Connect"), Guid("78A97F88-B796-403D-9658-B379E5385512")]
	//[CustomUI("RibbonUI.xml", true)]
    //[CustomPane(typeof(InnerAddinPane), "InnerAddin Pane", true)]
    [RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser), Timestamp]
    [DontRegisterAddin]
    public class Connect : PowerPoint.Tools.COMAddin
	{
		public Connect()
		{
            AppDomain.CurrentDomain.DomainUnload += CurrentDomain_DomainUnload;
        }

        protected override void RecieveCustomData(string custom)
        {
            string message = null != custom ? custom : "NULL";
            MessageBox.Show("I recieve custom data from an update/reload handler: " + message, "RecieveCustomData");
        }

        protected override void RaiseOnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            //MessageBox.Show("RaiseOnConnection Application: " + System.ComponentModel.TypeDescriptor.GetClassName(application));
            PrintArray("RaiseOnConnection", custom);
            base.RaiseOnConnection(application, connectMode, addInInst, ref custom);
        }

        protected override void RaiseOnAddInsUpdate(ref Array custom)
        {
            PrintArray("RaiseOnAddInsUpdate", custom);
            base.RaiseOnAddInsUpdate(ref custom);
        }

        protected override void RaiseOnStartupComplete(ref Array custom)
        {
            PrintArray("RaiseOnStartupComplete", custom);
            base.RaiseOnStartupComplete(ref custom);
        }

        private void CurrentDomain_DomainUnload(object sender, EventArgs e)
        {
            Trace.WriteLine("CurrentDomain_DomainUnload");
        }

        public bool HasShimHost
        {
            get
            {
                bool b = false;
                if (null != ShimHost)
                {
                    ShimHost.IsAvailable(ref b);
                }

                return null != ShimHost && true == b;
            }
        }

        public void CallShimHost()
        {
            //MessageBox.Show("CallShimHost");

            foreach (var item in TaskPanes.ToArray())
            {
                if (null != item.Pane)
                    item.Pane.Delete();
            }

            if (null != ShimHost)
            {
                ShimHost.Update("Greetings from previous addin instance.");
            }
            else
            {
                MessageBox.Show("I dont have a shim host :(");
            }
        }

        private void PrintArray(string caller, Array custom)
        {
            string message = caller;
            if (null != custom)
            {
                foreach (var item in custom)
                {
                    message += "," + (null != item ? item.ToString() : "<Empty>");
                }
            }
            else
            {
                message += "<Empty>";
            }
            Trace.WriteLine(message);
        }

        protected override bool QueryInterface(Guid interfaceId, ref Type type, ref object instance)
        {
            var iids = new NetOffice.ComTypes.WellKnownIID();
            Trace.WriteLine("QueryInterface " + iids.GetIID(interfaceId));

            if(iids.IID_IRibbonExtensibility == interfaceId)
            {
                type = typeof(NetOffice.OfficeApi.Native.IRibbonExtensibility);
                instance = new MyRibbonExtensibility(this);
                return true;
            }
            else if (iids.IID_ICustomTaskPaneConsumer == interfaceId)
            {
                type = typeof(NetOffice.OfficeApi.Native.ICustomTaskPaneConsumer);
                instance = new MyCustomTaskPaneConsumer(this);
                return true;
            }
            else
            {
                return base.QueryInterface(interfaceId, ref type, ref instance);
            }
        }

        protected override void OnError(ErrorMethodKind methodKind, System.Exception exception)
        {
            Trace.WriteLine("OnError " + methodKind.ToString() + " " + exception.Message);
            MessageBox.Show(exception.ToString(), methodKind.ToString());
        }

        [RegisterErrorHandler]
		public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, System.Exception exception)
		{
			Office.Tools.Contribution.DialogUtils.ShowRegisterError("InnerAddin", methodKind, exception);
		}
    }
}
