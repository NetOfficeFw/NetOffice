using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    #region Comments
    /// <summary>
    /// A utility class that exposes named pipes operations.
    /// </summary>
    /// <remarks>
    /// This class uses the exposed exposed kernel32.dll methods by the 
    /// <see cref="AppModule.NamedPipes.NamedPipeNative">NamedPipeNative</see> class
    /// to provided controlled named pipe functionality.
    /// </remarks>
    #endregion
    public sealed class NamedPipeWrapper
    {
        #region Comments
        /// <summary>
        /// The number of retries when creating a pipe or connecting to a pipe.
        /// </summary>
        #endregion
        private const int ATTEMPTS = 2;
        #region Comments
        /// <summary>
        /// Wait time for the 
        /// <see cref="AppModule.NamedPipes.NamedPipeNative.WaitNamedPipe">NamedPipeNative.WaitNamedPipe</see> 
        /// operation.
        /// </summary>
        #endregion
        private const int WAIT_TIME = 5000;
        #region Comments
        /// <summary>
        /// Reads a string from a named pipe using the UTF8 encoding.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        /// <param name="maxBytes">The maximum bytes to read.</param>
        /// <returns>A UTF8 string.</returns>
        /// <remarks>This function uses 
        /// <see cref="AppModule.NamedPipes.NamedPipeWrapper.ReadBytes">AppModule.NamedPipes.ReadBytes</see> 
        /// to read the bytes from the pipe and then converts them to string.<br/><br/>
        /// The first four bytes of the pipe data are expected to contain 
        /// the data length of the message. This method first reads those four 
        /// bytes and converts them to integer. It then continues to read from the pipe using 
        /// the extracted data length.
        /// </remarks>
        #endregion
        public static string Read(PipeHandle handle, int maxBytes)
        {
            string returnVal = "";
            byte[] bytes = ReadBytes(handle, maxBytes);
            if (bytes != null)
            {
                returnVal = System.Text.Encoding.UTF8.GetString(bytes);
            }
            return returnVal;
        }
        #region Comments
        /// <summary>
        /// Reads the bytes from a named pipe.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        /// <param name="maxBytes">The maximum bytes to read.</param>
        /// <returns>An array of bytes.</returns>
        /// <remarks>This method expects that the first four bytes in the pipe define 
        /// the length of the data to read. If the data length is greater than 
        /// <b>maxBytes</b> the method returns null.<br/><br/>
        /// The first four bytes of the pipe data are expected to contain 
        /// the data length of the message. This method first reads those four 
        /// bytes and converts them to integer. It then continues to read from the pipe using 
        /// the extracted data length.
        /// </remarks>
        #endregion
        public static byte[] ReadBytes(PipeHandle handle, int maxBytes)
        {
            byte[] numReadWritten = new byte[4];
            byte[] intBytes = new byte[4];
            byte[] msgBytes = null;
            int len;

            // Set the Handle state to Reading
            handle.State = InterProcessConnectionState.Reading;
            // Read the first four bytes and convert them to integer
            if (NamedPipeNative.ReadFile(handle.Handle, intBytes, 4, numReadWritten, 0))
            {
                len = BitConverter.ToInt32(intBytes, 0);
                msgBytes = new byte[len];
                // Read the rest of the data
                if (!NamedPipeNative.ReadFile(handle.Handle, msgBytes, (uint)len, numReadWritten, 0))
                {
                    handle.State = InterProcessConnectionState.Error;
                    return null;
//                    throw new NamedPipeIOException("Error reading from pipe. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
                }
            }
            else
            {
                handle.State = InterProcessConnectionState.Error;
                return null;
                //throw new NamedPipeIOException("Error reading from pipe. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
            }
            handle.State = InterProcessConnectionState.ReadData;
            if (len > maxBytes)
            {
                return null;
            }
            return msgBytes;
        }
        #region Comments
        /// <summary>
        /// Writes a string to a named pipe.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        /// <param name="text">The text to write to the pipe.</param>
        /// <remarks>This method converts the text into an array of bytes, using the 
        /// UTF8 encoding and the uses 
        /// <see cref="AppModule.NamedPipes.NamedPipeWrapper.WriteBytes">AppModule.NamedPipes.WriteBytes</see>
        /// to write to the pipe.<br/><br/>
        /// When writing to a pipe the method first writes four bytes that define the data length.
        /// It then writes the whole message.</remarks>
        #endregion
        public static void Write(PipeHandle handle, string text)
        {
            WriteBytes(handle, System.Text.Encoding.UTF8.GetBytes(text));
        }
        #region Comments
        /// <summary>
        /// Writes an array of bytes to a named pipe.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        /// <param name="bytes">The bytes to write.</param>
        /// <remarks>If we try bytes array we attempt to write is empty then this method write a space character to the pipe. This is necessary because the other end of the pipe uses a blocking Read operation so we must write someting.<br/><br/>
        /// The bytes length is restricted by the <b>maxBytes</b> parameter, which is done primarily for security reasons.<br/><br/>
        /// When writing to a pipe the method first writes four bytes that define the data length.
        /// It then writes the whole message.</remarks>
        #endregion
        public static void WriteBytes(PipeHandle handle, byte[] bytes)
        {
            byte[] numReadWritten = new byte[4];
            uint len;

            if (bytes == null)
            {
                bytes = new byte[0];
            }
            if (bytes.Length == 0)
            {
                bytes = new byte[1];
                bytes = System.Text.Encoding.UTF8.GetBytes(" ");
            }
            // Get the message length
            len = (uint)bytes.Length;
            handle.State = InterProcessConnectionState.Writing;
            // Write four bytes that define the message length
            if (NamedPipeNative.WriteFile(handle.Handle, BitConverter.GetBytes(len), 4, numReadWritten, 0))
            {
                // Write the whole message
                if (!NamedPipeNative.WriteFile(handle.Handle, bytes, len, numReadWritten, 0))
                {
                    handle.State = InterProcessConnectionState.Error;
                    throw new NamedPipeIOException("Error writing to pipe. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
                }
            }
            else
            {
                handle.State = InterProcessConnectionState.Error;
                throw new NamedPipeIOException("Error writing to pipe. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
            }
            handle.State = InterProcessConnectionState.Flushing;
            Flush(handle);
            handle.State = InterProcessConnectionState.FlushedData;
        }
        #region Comments
        /// <summary>
        /// Tries to connect to a named pipe on the same machine.
        /// </summary>
        /// <param name="pipeName">The name of the pipe.</param>
        /// <param name="handle">The resulting pipe handle.</param>
        /// <returns>Return true if the attempt succeeds.</returns>
        /// <remarks>This method is used mainly when stopping the pipe server. It unblocks the existing pipes, which wait for client connection.</remarks>
        #endregion
        public static bool TryConnectToPipe(string pipeName, out PipeHandle handle)
        {
            return TryConnectToPipe(pipeName, ".", out handle);
        }
        #region Comments
        /// <summary>
        /// Tries to connect to a named pipe.
        /// </summary>
        /// <param name="pipeName">The name of the pipe.</param>
        /// <param name="serverName">The name of the server.</param>
        /// <param name="handle">The resulting pipe handle.</param>
        /// <returns>Return true if the attempt succeeds.</returns>
        /// <remarks>This method is used mainly when stopping the pipe server. It unblocks the existing pipes, which wait for client connection.</remarks>
        #endregion
        public static bool TryConnectToPipe(string pipeName, string serverName, out PipeHandle handle)
        {
            handle = new PipeHandle();
            // Build the pipe name string
            string name = @"\\" + serverName + @"\pipe\" + pipeName;
            handle.State = InterProcessConnectionState.ConnectingToServer;
            // Try to connect to a server pipe
            handle.Handle = NamedPipeNative.CreateFile(name, NamedPipeNative.GENERIC_READ | NamedPipeNative.GENERIC_WRITE, 0, null, NamedPipeNative.OPEN_EXISTING, 0, 0);
            if (handle.Handle.ToInt32() != NamedPipeNative.INVALID_HANDLE_VALUE)
            {
                handle.State = InterProcessConnectionState.ConnectedToServer;
                return true;
            }
            else
            {
                handle.State = InterProcessConnectionState.Error;
                return false;
            }
        }
        #region Comments
        /// <summary>
        /// Connects to a server named pipe on the same machine.
        /// </summary>
        /// <param name="pipeName">The pipe name.</param>
        /// <returns>The pipe handle, which also contains the pipe state.</returns>
        /// <remarks>This method is used by clients to establish a pipe connection with a server pipe.</remarks>
        #endregion
        public static PipeHandle ConnectToPipe(string pipeName)
        {
            return ConnectToPipe(pipeName, ".");
        }
        #region Comments
        /// <summary>
        /// Connects to a server named pipe.
        /// </summary>
        /// <param name="pipeName">The pipe name.</param>
        /// <param name="serverName">The server name.</param>
        /// <returns>The pipe handle, which also contains the pipe state.</returns>
        /// <remarks>This method is used by clients to establish a pipe connection with a server pipe.</remarks>
        #endregion
        public static PipeHandle ConnectToPipe(string pipeName, string serverName)
        {
            PipeHandle handle = new PipeHandle();
            // Build the name of the pipe.
            string name = @"\\" + serverName + @"\pipe\" + pipeName;

            for (int i = 1; i <= ATTEMPTS; i++)
            {
                handle.State = InterProcessConnectionState.ConnectingToServer;
                // Try to connect to the server
                handle.Handle = NamedPipeNative.CreateFile(name, NamedPipeNative.GENERIC_READ | NamedPipeNative.GENERIC_WRITE, 0, null, NamedPipeNative.OPEN_EXISTING, 0, 0);
                if (handle.Handle.ToInt32() != NamedPipeNative.INVALID_HANDLE_VALUE)
                {
                    // The client managed to connect to the server pipe
                    handle.State = InterProcessConnectionState.ConnectedToServer;
                    // Set the read mode of the pipe channel
                    uint mode = NamedPipeNative.PIPE_READMODE_MESSAGE;
                    if (NamedPipeNative.SetNamedPipeHandleState(handle.Handle, ref mode, IntPtr.Zero, IntPtr.Zero))
                    {
                        break;
                    }
                    if (i >= ATTEMPTS)
                    {
                        handle.State = InterProcessConnectionState.Error;
                        throw new NamedPipeIOException("Error setting read mode on pipe " + name + " . Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
                    }
                }
                if (i >= ATTEMPTS)
                {
                    if (NamedPipeNative.GetLastError() != NamedPipeNative.ERROR_PIPE_BUSY)
                    {
                        handle.State = InterProcessConnectionState.Error;
                        // After a certain number of unsuccessful attempt raise an exception
                        throw new NamedPipeIOException("Error connecting to pipe " + name + " . Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
                    }
                    else
                    {
                        handle.State = InterProcessConnectionState.Error;
                        throw new NamedPipeIOException("Pipe " + name + " is too busy. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
                    }
                }
                else
                {
                    // The pipe is busy so lets wait for some time and try again
                    if (NamedPipeNative.GetLastError() == NamedPipeNative.ERROR_PIPE_BUSY)
                        NamedPipeNative.WaitNamedPipe(name, WAIT_TIME);
                }
            }

            return handle;
        }
        #region Comments
        /// <summary>
        /// Creates a server named pipe.
        /// </summary>
        /// <param name="name">The name of the pipe.</param>
        /// <param name="outBuffer">The size of the outbound buffer.</param>
        /// <param name="inBuffer">The size of the inbound buffer.</param>
        /// <returns>The pipe handle.</returns>
        #endregion
        public static PipeHandle Create(string name, uint outBuffer, uint inBuffer)
        {
            name = @"\\.\pipe\" + name;
            PipeHandle handle = new PipeHandle();
            for (int i = 1; i <= ATTEMPTS; i++)
            {
                handle.State = InterProcessConnectionState.Creating;
                handle.Handle = NamedPipeNative.CreateNamedPipe(
                    name,
                    NamedPipeNative.PIPE_ACCESS_DUPLEX,
                    NamedPipeNative.PIPE_TYPE_MESSAGE | NamedPipeNative.PIPE_READMODE_MESSAGE | NamedPipeNative.PIPE_WAIT,
                    NamedPipeNative.PIPE_UNLIMITED_INSTANCES,
                    outBuffer,
                    inBuffer,
                    NamedPipeNative.NMPWAIT_WAIT_FOREVER,
                    IntPtr.Zero);
                if (handle.Handle.ToInt32() != NamedPipeNative.INVALID_HANDLE_VALUE)
                {
                    handle.State = InterProcessConnectionState.Created;
                    break;
                }
                if (i >= ATTEMPTS)
                {
                    handle.State = InterProcessConnectionState.Error;
                    throw new NamedPipeIOException("Error creating named pipe " + name + " . Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
                }
            }

            return handle;
        }
        #region Comments
        /// <summary>
        /// Starts waiting for client connections.
        /// </summary>
        /// <remarks>
        /// Blocks the current execution until a client pipe attempts to establish a connection.
        /// </remarks>
        /// <param name="handle">The pipe handle.</param>
        #endregion
        public static void Connect(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.WaitingForClient;
            bool connected = NamedPipeNative.ConnectNamedPipe(handle.Handle, null);
            handle.State = InterProcessConnectionState.ConnectedToClient;
            if (!connected && NamedPipeNative.GetLastError() != NamedPipeNative.ERROR_PIPE_CONNECTED)
            {
                handle.State = InterProcessConnectionState.Error;
                throw new NamedPipeIOException("Error connecting pipe. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
            }
        }
        #region Comments
        /// <summary>
        /// Returns the number of instances of a named pipe.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        /// <returns>The number of instances.</returns>
        #endregion
        public static uint NumberPipeInstances(PipeHandle handle)
        {
            uint curInstances = 0;

            if (NamedPipeNative.GetNamedPipeHandleState(handle.Handle, IntPtr.Zero, ref curInstances, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero))
            {
                return curInstances;
            }
            else
            {
                throw new NamedPipeIOException("Error getting the pipe state. Internal error: " + NamedPipeNative.GetLastError().ToString(), NamedPipeNative.GetLastError());
            }
        }
        #region Comments
        /// <summary>
        /// Closes a named pipe and releases the native handle.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        #endregion
        public static void Close(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.Closing;
            NamedPipeNative.CloseHandle(handle.Handle);
            handle.Handle = IntPtr.Zero;
            handle.State = InterProcessConnectionState.Closed;
        }
        #region Comments
        /// <summary>
        /// Flushes all the data in a named pipe.
        /// </summary>
        /// <param name="handle">The pipe handle.</param>
        #endregion
        public static void Flush(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.Flushing;
            NamedPipeNative.FlushFileBuffers(handle.Handle);
            handle.State = InterProcessConnectionState.FlushedData;
        }
        #region Comments
        /// <summary>
        /// Disconnects a server named pipe from the client.
        /// </summary>
        /// <remarks>
        /// Server pipes can be reused by first disconnecting them from the client and then
        /// calling the <see cref="AppModule.NamedPipes.NamedPipeWrapper.Connect">Connect</see>
        /// method to start listening. This improves the performance as it is not necessary
        /// to create new pipe handles.
        /// </remarks>
        /// <param name="handle">The pipe handle.</param>
        #endregion
        public static void Disconnect(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.Disconnecting;
            NamedPipeNative.DisconnectNamedPipe(handle.Handle);
            handle.State = InterProcessConnectionState.Disconnected;
        }
        #region Comments
        /// <summary>
        /// Private constructor.
        /// </summary>
        #endregion
        private NamedPipeWrapper() { }
    }
}
