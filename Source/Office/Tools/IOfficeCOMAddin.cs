using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    ///  Represents an addin implementation trough IDTExtensibility2 and Ribbon/TaskPane support
    /// </summary>
    public interface IOfficeCOMAddin : NetOffice.Tools.ICOMAddin, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
    }
}
