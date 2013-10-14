using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    #region Comments
    /// <summary>
    /// Used by client applications to communicate with server ones by using named pipes.
    /// </summary>
    #endregion
    public sealed class ClientPipeConnection : APipeConnection
    {
        #region Comments
        /// <summary>
        /// The network name of the server where the server pipe is created.
        /// </summary>
        /// <remarks>
        /// If "." is used as a server name then the pipe is connected to the local machine.
        /// </remarks>
        #endregion
        private string Server = ".";
        #region Comments
        /// <summary>
        /// Closes a client named pipe connection.
        /// </summary>
        /// <remarks>
        /// A client pipe connection is closed by closing the underlying pipe handle.
        /// </remarks>
        #endregion
        public override void Close()
        {
            CheckIfDisposed();
            NamedPipeWrapper.Close(this.Handle);
        }
        #region Comments
        /// <summary>
        /// Connects a client pipe to an existing server one.
        /// </summary>
        #endregion
        public override void Connect()
        {
            CheckIfDisposed();
            this.Handle = NamedPipeWrapper.ConnectToPipe(this.Name, this.Server);
        }
        #region Comments
        /// <summary>
        /// Attempts to establish a connection to the a server named pipe. 
        /// </summary>
        /// <remarks>
        /// If the attempt is successful the method creates the 
        /// <see cref="AppModule.NamedPipes.PipeHandle">PipeHandle</see> object
        /// and assigns it to the <see cref="AppModule.NamedPipes.APipeConnection.Handle">Handle</see>
        /// field.<br/><br/>
        /// This method is used when it is not known whether a server pipe already exists.
        /// </remarks>
        /// <returns>True if a connection is established.</returns>
        #endregion
        public bool TryConnect()
        {
            CheckIfDisposed();
            bool ReturnVal = NamedPipeWrapper.TryConnectToPipe(this.Name, this.Server, out this.Handle);

            return ReturnVal;
        }
        #region Comments
        /// <summary>
        /// Creates an instance of the ClientPipeConnection assuming that the server pipe
        /// is created on the same machine.
        /// </summary>
        /// <remarks>
        /// The maximum bytes to read from the client is set to be Int32.MaxValue.
        /// </remarks>
        /// <param name="name">The name of the server pipe.</param>
        #endregion
        public ClientPipeConnection(string name)
        {
            this.Name = name;
            this.Server = ".";
            this.maxReadBytes = Int32.MaxValue;
        }
        #region Comments
        /// <summary>
        /// Creates an instance of the ClientPipeConnection specifying the network name
        /// of the server.
        /// </summary>
        /// <remarks>
        /// The maximum bytes to read from the client is set to be Int32.MaxValue.
        /// </remarks>
        /// <param name="name">The name of the server pipe.</param>
        /// <param name="server">The network name of the machine, where the server pipe is created.</param>
        #endregion
        public ClientPipeConnection(string name, string server)
        {
            this.Name = name;
            this.Server = server;
            this.maxReadBytes = Int32.MaxValue;
        }
        #region Comments
        /// <summary>
        /// Object destructor.
        /// </summary>
        #endregion
        ~ClientPipeConnection()
        {
            Dispose(false);
        }
    }
}
