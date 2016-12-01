using System;
using System.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Utils
{
    /// <summary>
    /// Provides arguments for the KeyPress event.
    /// </summary>
    public class TrayMenuItemKeyPressEventArgs : EventArgs
    {
        /// <summary>
        /// Creates a new instance of the class
        /// </summary>
        /// <param name="keyChar">The ASCII character corresponding to the key the user pressed.</param>
        [TargetedPatchingOptOut("Performance critical to inline this type of method across NGen image boundaries")]
        public TrayMenuItemKeyPressEventArgs(char keyChar)
        {
            KeyChar = keyChar;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the KeyPress event was handled.
        /// </summary>
        public bool Handled { get; set; }

        /// <summary>
        /// Gets or sets the character corresponding to the key pressed.
        /// </summary>
        public char KeyChar { get; set; }
    }

    /// <summary>
    /// Key Press Event Handler
    /// </summary>
    /// <param name="sender">sender instance</param>
    /// <param name="args">key press arguments</param>
    public delegate void TrayMenuItemKeyPressEventHandler(object sender, TrayMenuItemKeyPressEventArgs args);
}
