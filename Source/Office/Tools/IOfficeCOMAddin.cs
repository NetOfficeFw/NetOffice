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
    public interface IOfficeCOMAddin : NetOffice.Tools.ICOMAddin, Native.IRibbonExtensibility, Native.ICustomTaskPaneConsumer
    {
        /// <summary>
        /// Host Application
        /// </summary>
        ICOMObject Application { get; }

        /// <summary>
        /// Used Factory Core
        /// </summary>
        Core Factory { get; }

        /// <summary>
        /// Instance Type (cached)
        /// </summary>
        Type Type { get; }

        /// <summary>
        /// Collection with all created custom Task Panes
        /// </summary>
        CustomTaskPaneCollection TaskPanes { get; }
    }
}
