using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.NamedPipes
{
    internal sealed class NamedPipeWrapper
    {
        private const int ATTEMPTS = 2;
        private const int WAIT_TIME = 5000;

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
                }
            }
            else
            {
                handle.State = InterProcessConnectionState.Error;
                return null;
            }
            handle.State = InterProcessConnectionState.ReadData;
            if (len > maxBytes)
            {
                return null;
            }
            return msgBytes;
        }

        public static void Write(PipeHandle handle, string text)
        {
            WriteBytes(handle, System.Text.Encoding.UTF8.GetBytes(text));
        }

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
                    return;
                }
            }
            else
            {
                handle.State = InterProcessConnectionState.Error;
                return;
            }
            handle.State = InterProcessConnectionState.Flushing;
            Flush(handle);
            handle.State = InterProcessConnectionState.FlushedData;
        }

        public static bool TryConnectToPipe(string pipeName, out PipeHandle handle)
        {
            return TryConnectToPipe(pipeName, ".", out handle);
        }

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

        public static PipeHandle ConnectToPipe(string pipeName)
        {
            return ConnectToPipe(pipeName, ".");
        }

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
                        return null;
                    }
                }
                if (i >= ATTEMPTS)
                {
                    if (NamedPipeNative.GetLastError() != NamedPipeNative.ERROR_PIPE_BUSY)
                    {
                        handle.State = InterProcessConnectionState.Error;
                        // After a certain number of unsuccessful attempt raise an exception
                        return null;
                    }
                    else
                    {
                        handle.State = InterProcessConnectionState.Error;
                        return null;
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
                    return null;
                }
            }

            return handle;
        }

        public static void Connect(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.WaitingForClient;
            bool connected = NamedPipeNative.ConnectNamedPipe(handle.Handle, null);
            handle.State = InterProcessConnectionState.ConnectedToClient;
            if (!connected && NamedPipeNative.GetLastError() != NamedPipeNative.ERROR_PIPE_CONNECTED)
            {
                handle.State = InterProcessConnectionState.Error;
            }
        }

        public static uint NumberPipeInstances(PipeHandle handle)
        {
            uint curInstances = 0;

            if (NamedPipeNative.GetNamedPipeHandleState(handle.Handle, IntPtr.Zero, ref curInstances, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero))
            {
                return curInstances;
            }
            else
            {
                throw new Exception("Error getting the pipe state. Internal error: " + NamedPipeNative.GetLastError().ToString());
            }
        }

        public static void Close(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.Closing;
            NamedPipeNative.CloseHandle(handle.Handle);
            handle.Handle = IntPtr.Zero;
            handle.State = InterProcessConnectionState.Closed;
        }

        public static void Flush(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.Flushing;
            NamedPipeNative.FlushFileBuffers(handle.Handle);
            handle.State = InterProcessConnectionState.FlushedData;
        }

        public static void Disconnect(PipeHandle handle)
        {
            handle.State = InterProcessConnectionState.Disconnecting;
            NamedPipeNative.DisconnectNamedPipe(handle.Handle);
            handle.State = InterProcessConnectionState.Disconnected;
        }

        private NamedPipeWrapper() { }
    }
}
