using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;

namespace ExampleBase
{
    /// <summary>
    /// the primary interface for an example
    /// </summary>
    public interface IExample
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
        /// Visual panel from the example, can be null
        /// </summary>
        UserControl Panel { get; }

        /// <summary>
        /// called from IHost after construction
        /// </summary>
        /// <param name="hostApplication">the Host Application for the examples</param>
        void Connect(IHost hostApplication);

        /// <summary>
        /// Run the example
        /// </summary>
        void RunExample();
    }
}
