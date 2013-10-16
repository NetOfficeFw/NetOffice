using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice
{
    /// <summary>
    /// event args for OnDispose
    /// </summary>
    public class OnDisposeEventArgs : EventArgs
    {
        /// <summary>
        /// creates a new instance
        /// </summary>
        /// <param name="sender">the target COM object</param>
        internal OnDisposeEventArgs(COMObject sender)
        {
            Sender = sender;
        }

        /// <summary>
        /// The target COM object
        /// </summary>
        public COMObject Sender { get; private set; }

        /// <summary>
        /// Skip flag, you can cancel the operation if you want
        /// </summary>
        public bool Cancel { get; set; }
    }

    /// <summary>
    /// EventHandler delegate for COMObject.OnDispose
    /// </summary>
    /// <param name="eventArgs">info about the target instance</param>
    public delegate void OnDisposeEventHandler(OnDisposeEventArgs eventArgs);
}
