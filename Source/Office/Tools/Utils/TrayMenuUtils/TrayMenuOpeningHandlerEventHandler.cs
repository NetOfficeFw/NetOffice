using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Tray menu opening event handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="args">cancel arguments</param>
    public delegate void TrayMenuOpeningHandlerEventHandler(object sender, CancelEventArgs args);
}
