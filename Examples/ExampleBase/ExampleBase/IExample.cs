using System;
using System.Windows.Forms;

namespace ExampleBase
{
    /// <summary>
    /// Represents a single example
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
        /// Visual panel from the example, can be null(Nothing in Visual Basic)
        /// </summary>
        UserControl Panel { get; }
        
        /// <summary>
        /// Called from IHost while connecting to host application
        /// </summary>
        /// <param name="hostApplication">the host application for the example</param>
        void Connect(IHost hostApplication);

        /// <summary>
        /// Run the example
        /// </summary>
        void RunExample();
    }
}