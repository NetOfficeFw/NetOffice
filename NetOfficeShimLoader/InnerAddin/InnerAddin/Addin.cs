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

namespace InnerAddin
{
	//[COMAddin("InnerAddin", "InnerAddin Description", LoadBehavior.LoadAtStartup), ProgId("InnerAddin.Addin"), Guid("78A97F88-B796-403D-9658-B379E5385512")]
	[CustomUI("RibbonUI.xml", true)]
    [CustomPane(typeof(InnerAddinPane), "InnerAddin Pane", true)]
    [RegistryLocation(RegistrySaveLocation.InstallScopeCurrentUser)]
    [DontRegisterAddin]
    public class Addin : PowerPoint.Tools.COMAddin
	{
		public Addin()
		{
            this.OnBeginShutdown += Addin_OnBeginShutdown;
            this.OnConnection += Addin_OnConnection;
			this.OnStartupComplete += Addin_OnStartupComplete;
			this.OnDisconnection += Addin_OnDisconnection;
            this.OnAddInsUpdate += Addin_OnAddInsUpdate;
        }

        protected override bool QueryInterface(Guid interfaceId, ref Type type, ref object instance)
        {
            var iids = new NetOffice.ComTypes.WellKnownIID();
            //MessageBox.Show("QueryInterface " + iids.GetIID(interfaceId));

            if(iids.IID_IRibbonExtensibility == interfaceId)
            {
                type = typeof(NetOffice.OfficeApi.Native.IRibbonExtensibility);
                instance = new MyRibbonExtensibility();
                return true;
            }

            if (interfaceId == Guid.Parse("e19c7100-9709-4db7-9373-e7b518b47086"))
            {
                //MessageBox.Show("QueryInterface decline " + iids.GetIID(interfaceId));
                return true;
            }
            else
            {
                return base.QueryInterface(interfaceId, ref type, ref instance);
            }
        }

        private void Addin_OnAddInsUpdate(ref Array custom)
        {
            Console.WriteLine("Addin_OnAddInsUpdate");
        }

        private void Addin_OnBeginShutdown(ref Array custom)
        {
            Console.WriteLine("Addin_OnBeginShutdown");
        }

        private void Addin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Console.WriteLine("Addin_OnConnection");
            //MessageBox.Show("InnerAddin Addin_OnConnection");
        }

        private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Console.WriteLine("Addin_OnDisconnection");
        }

        private void Addin_OnStartupComplete(ref Array custom)
		{
            Console.WriteLine("Addin_OnStartupComplete");
            //MessageBox.Show("Addin_OnStartupComplete");
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
