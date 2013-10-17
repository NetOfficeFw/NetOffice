using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace TutorialsBase
{
    /// <summary>
    /// the primary interface for a tutorial
    /// </summary>
    public interface ITutorial
    {
        /// <summary>
        /// Friendly name of the example
        /// </summary>
        string Caption { get; }

        /// <summary>
        /// Description of the example
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Path the HTML documentation page
        /// </summary>
        string Uri { get; }

        /// <summary>
        /// The language in the host application was changed
        /// </summary>
        /// <param name="lcid">1031(german) or 1033(english)</param>
        void ChangeLanguage(int lcid);

        /// <summary>
        /// Visual panel from the example
        /// </summary>
        UserControl Panel { get; }

        /// <summary>
        /// called from IHost after construction
        /// </summary>
        /// <param name="hostApplication">the Host Application for the examples</param>
        void Connect(IHost hostApplication);

        /// <summary>
        /// called from IHost when the application is shutdown
        /// </summary>
        void Disconnect();

        /// <summary>
        /// The host application shows a Run button for tutorials without an own Panel. this method was called if anybody clicks on the button
        /// </summary>
        void Run();
    }
}
