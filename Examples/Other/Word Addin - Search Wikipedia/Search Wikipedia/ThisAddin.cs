using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;

using NetOffice.Tools;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi.Tools;

namespace Sample.Addin
{
    /// <summary>
    /// The main addin for MS-Word. The Addin use the base class COMAddin from NetOffice.WordApi.Tools.
    /// Learn more about the NetOffice Tools namespace: http://netoffice.codeplex.com/wikipage?title=Tools_EN
    /// </summary>
    [GuidAttribute("56F843AD-ECB8-45D6-9E33-C0928BD2FB03"), ProgId("Sample.WordAddin"), ComVisible(true)]
    [COMAddin("Word Wikipedia Addin", "This Addin provides Wikipedia functionality", 3)]
    public class ThisAddin : COMAddin
    {
        public ThisAddin()
        {
            // we create the taskpane
            TaskPanes.Add(typeof(WikipediaPane), "Wikipedia - NetOffice Sample");
            TaskPanes[0].DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            TaskPanes[0].DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
            TaskPanes[0].Width = 520;
            TaskPanes[0].Visible = true;
        }
    }
}
