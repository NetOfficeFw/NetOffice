using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Tools;
using NetOffice.Tools;
using System.ComponentModel;

namespace InnerAddin
{
    [CustomPane(typeof(InnerAddinPane), "InnerAddin Pane", true)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class MyCustomTaskPaneConsumer : NetOffice.OfficeApi.Tools.CustomTaskPaneConsumer
    {
        public MyCustomTaskPaneConsumer(Addin parent) : base(parent)
        {
            
        }



        protected override bool QueryInterface(Guid interfaceId, ref Type type, ref object instance)
        {
            NetOffice.ComTypes.WellKnownIID id = new NetOffice.ComTypes.WellKnownIID();
            MessageBox.Show("Query " + id.GetIID(interfaceId));
            return base.QueryInterface(interfaceId, ref type, ref instance);
        }
    }
}
