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

namespace InnerAddin
{
	[COMAddin("InnerAddin", "InnerAddin Description", LoadBehavior.LoadAtStartup), ProgId("InnerAddin.Addin"), Guid("78A97F88-B796-403D-9658-B379E5385512")]
	//[CustomUI("RibbonUI.xml", true)]
    //[CustomPane(typeof(InnerAddinPane), "InnerAddin Pane", true)]
    [RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser), Timestamp]
    [DontRegisterAddin]
    public class Addin : PowerPoint.Tools.COMAddin
	{
		public Addin()
		{
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;
            AppDomain.CurrentDomain.TypeResolve += CurrentDomain_TypeResolve;
            bool b = AppDomain.CurrentDomain.IsDefaultAppDomain();
            if (!b)
                MessageBox.Show("no default domain");
        }

        private Assembly CurrentDomain_TypeResolve(object sender, ResolveEventArgs args)
        {
            MessageBox.Show("CurrentDomain_TypeResolve " + args.Name);
            return null;
        }

        private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            MessageBox.Show("CurrentDomain_AssemblyResolve " + args.Name);
            return null;
        }

        public bool HasShimHost
        {
            get
            {
                return null != ShimHost;
            }
        }

        public void CallShimHost()
        {
            foreach (var item in TaskPanes.ToArray())
            {
                if(null != item.Pane)
                    item.Pane.Delete();
            }

            if (null != ShimHost)
                ShimHost.Reload();
        }

        protected override bool QueryInterface(Guid interfaceId, ref Type type, ref object instance)
        {
            var iids = new NetOffice.ComTypes.WellKnownIID();
            //MessageBox.Show("QueryInterface " + iids.GetIID(interfaceId));

            if(iids.IID_IRibbonExtensibility == interfaceId)
            {
                type = typeof(NetOffice.OfficeApi.Native.IRibbonExtensibility);
                instance = new MyRibbonExtensibility(this);
                return true;
            }
            else if (iids.IID_ICustomTaskPaneConsumer == interfaceId)
            {
                //MessageBox.Show("QueryInterface " + iids.GetIID(interfaceId));

                type = typeof(NetOffice.OfficeApi.Native.ICustomTaskPaneConsumer);
                instance = new MyCustomTaskPaneConsumer(this);
                return true;
            }
            else if (interfaceId == Guid.Parse("e19c7100-9709-4db7-9373-e7b518b47086"))
            {
                //MessageBox.Show("QueryInterface decline " + iids.GetIID(interfaceId));
                return true;
            }
            else
            {
                return base.QueryInterface(interfaceId, ref type, ref instance);
            }
        }

        protected override void OnError(ErrorMethodKind methodKind, System.Exception exception)
		{
            MessageBox.Show(exception.ToString(), methodKind.ToString());
		}

        public void SampleButton_Click(Office.IRibbonControl control)
        {
            MessageBox.Show("Thanks!", "InnerAddin.Addin");
        }

        [RegisterErrorHandler]
		public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, System.Exception exception)
		{
			Office.Tools.Contribution.DialogUtils.ShowRegisterError("InnerAddin", methodKind, exception);
		}
    }
}
