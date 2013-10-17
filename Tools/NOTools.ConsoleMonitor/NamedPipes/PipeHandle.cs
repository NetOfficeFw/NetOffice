using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    #region Comments
    /// <summary>
    /// Holds the operating system native handle and the current state of the pipe connection.
    /// </summary>
    #endregion
    public sealed class PipeHandle
    {
        #region Comments
        /// <summary>
        /// The operating system native handle.
        /// </summary>
        #endregion
        public IntPtr Handle;
        #region Comments
        /// <summary>
        /// The current state of the pipe connection.
        /// </summary>
        #endregion
        public InterProcessConnectionState State;
        #region Comments
        /// <summary>
        /// Creates a PipeHandle instance using the passed native handle.
        /// </summary>
        /// <param name="hnd">The native handle.</param>
        #endregion
        public PipeHandle(int hnd)
        {
            this.Handle = new IntPtr(hnd);
            this.State = InterProcessConnectionState.NotSet;
        }
        #region Comments
        /// <summary>
        /// Creates a PipeHandle instance using the provided native handle and state.
        /// </summary>
        /// <param name="hnd">The native handle.</param>
        /// <param name="state">The state of the pipe connection.</param>
        #endregion
        public PipeHandle(int hnd, InterProcessConnectionState state)
        {
            this.Handle = new IntPtr(hnd);
            this.State = state;
        }
        #region Comments
        /// <summary>
        /// Creates a PipeHandle instance with an invalid native handle.
        /// </summary>
        #endregion
        public PipeHandle()
        {
            this.Handle = new IntPtr(NamedPipeNative.INVALID_HANDLE_VALUE);
            this.State = InterProcessConnectionState.NotSet;
        }
    }
}
