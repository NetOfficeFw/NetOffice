using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Represents the method that will handle the MouseDown, MouseUp, or MouseMove event of a form, control, or other component.    
    /// </summary>
    /// <param name="sender">The source of the event.</param>
    /// <param name="args">Arguments that contains the event data</param>
    public delegate void ToolsMouseEventHandler(object sender, ToolsMouseEventArgs args);
}
