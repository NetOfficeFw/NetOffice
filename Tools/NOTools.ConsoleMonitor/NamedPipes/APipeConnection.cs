using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    #region Comments
    /// <summary>
    /// An abstract class, which defines the methods for creating named pipes 
    /// connections, reading and writing data.
    /// </summary>
    /// <remarks>
    /// This class is inherited by 
    /// <see cref="AppModule.NamedPipes.ClientPipeConnection">ClientPipeConnection</see> 
    /// and <see cref="AppModule.NamedPipes.ServerPipeConnection">ServerPipeConnection</see> 
    /// classes, used for client and server applications respectively, which communicate
    /// using NamesPipes.
    /// </remarks>
    #endregion
    public abstract class APipeConnection : IInterProcessConnection
    {
        #region Comments
        /// <summary>
        /// A <see cref="AppModule.NamedPipes.PipeHandle">PipeHandle</see> object containing
        /// the native pipe handle.
        /// </summary>
        #endregion
        protected PipeHandle Handle = new PipeHandle();
        #region Comments
        /// <summary>
        /// The name of the named pipe.
        /// </summary>
        /// <remarks>
        /// This name is used for creating a server pipe and connecting client ones to it.
        /// </remarks>
        #endregion
        protected string Name;
        #region Comments
        /// <summary>
        /// Boolean field used by the IDisposable implementation.
        /// </summary>
        #endregion
        protected bool disposed = false;
        #region Comments
        /// <summary>
        /// The maximum bytes that will be read from the pipe connection.
        /// </summary>
        /// <remarks>
        /// This field could be used if the maximum length of the client message
        /// is known and we want to implement some security, which prevents the
        /// server from reading larger messages.
        /// </remarks>
        #endregion
        protected int maxReadBytes;
        #region Comments
        /// <summary>
        /// Reads a message from the pipe connection and converts it to a string
        /// using the UTF8 encoding.
        /// </summary>
        /// <remarks>
        /// See the <see cref="AppModule.NamedPipes.NamedPipeWrapper.Read">NamedPipeWrapper.Read</see>
        /// method for an explanation of the message format.
        /// </remarks>
        /// <returns>The UTF8 encoded string representation of the data.</returns>
        #endregion

        public string Read()
        {
            CheckIfDisposed();
            return NamedPipeWrapper.Read(Handle, maxReadBytes);
        }
        #region Comments
        /// <summary>
        /// Reads a message from the pipe connection.
        /// </summary>
        /// <remarks>
        /// See the <see cref="AppModule.NamedPipes.NamedPipeWrapper.ReadBytes">NamedPipeWrapper.ReadBytes</see>
        /// method for an explanation of the message format.
        /// </remarks>
        /// <returns>The bytes read from the pipe connection.</returns>
        #endregion
        public byte[] ReadBytes()
        {
            CheckIfDisposed();
            return NamedPipeWrapper.ReadBytes(Handle, maxReadBytes);
        }
        #region Comments
        /// <summary>
        /// Writes a string to the pipe connection/
        /// </summary>
        /// <param name="text">The text to write.</param>
        #endregion
        public void Write(string text)
        {
            CheckIfDisposed();
            NamedPipeWrapper.Write(Handle, text);
        }
        #region Comments
        /// <summary>
        /// Writes an array of bytes to the pipe connection.
        /// </summary>
        /// <param name="bytes">The bytes array.</param>
        #endregion
        public void WriteBytes(byte[] bytes)
        {
            CheckIfDisposed();
            NamedPipeWrapper.WriteBytes(Handle, bytes);
        }
        #region Comments
        /// <summary>
        /// Closes the pipe connection.
        /// </summary>
        #endregion
        public abstract void Close();
        #region Comments
        /// <summary>
        /// Connects a pipe connection.
        /// </summary>
        #endregion
        public abstract void Connect();
        #region Comments
        /// <summary>
        /// Disposes a pipe connection by closing the underlying native handle.
        /// </summary>
        #endregion
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #region Comments
        /// <summary>
        /// Disposes a pipe connection by closing the underlying native handle.
        /// </summary>
        /// <param name="disposing">A boolean indicating how the method is called.</param>
        #endregion
        protected void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                NamedPipeWrapper.Close(this.Handle);
            }
            disposed = true;
        }
       
        #region Comments
        /// <summary>
        /// Checks if the pipe connection is disposed.
        /// </summary>
        /// <remarks>
        /// This check is done before performing any pipe operations.
        /// </remarks>
        #endregion
        public void CheckIfDisposed()
        {
            if (this.disposed)
            {
                throw new ObjectDisposedException("The Pipe Connection is disposed.");
            }
        }

        public bool IsDisposed
        {
            get { return this.disposed; }
        }

        #region Comments
        /// <summary>
        /// Gets the pipe connection state from the <see cref="AppModule.NamedPipes.PipeHandle">PipeHandle</see> 
        /// object.
        /// </summary>
        /// <returns>The pipe connection state.</returns>
        #endregion
        public InterProcessConnectionState GetState()
        {
            CheckIfDisposed();
            return this.Handle.State;
        }
        #region Comments
        /// <summary>
        /// Retrieved the operating system native handle for the pipe connection.
        /// </summary>
        #endregion
        public int NativeHandle
        {
            get
            {
                CheckIfDisposed();
                return (int)this.Handle.Handle;
            }
        }
    }
}
