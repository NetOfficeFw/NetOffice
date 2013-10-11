using System;
using System.Security;
using System.Runtime.InteropServices;

namespace NetOffice.NamedPipes
{
    [SuppressUnmanagedCodeSecurity]
    internal sealed class NamedPipeNative
    {
        public const uint PIPE_ACCESS_OUTBOUND = 0x00000002;
        public const uint PIPE_ACCESS_DUPLEX = 0x00000003;
        public const uint PIPE_ACCESS_INBOUND = 0x00000001;
        public const uint PIPE_WAIT = 0x00000000;
        public const uint PIPE_NOWAIT = 0x00000001;
        public const uint PIPE_READMODE_BYTE = 0x00000000;
        public const uint PIPE_READMODE_MESSAGE = 0x00000002;
        public const uint PIPE_TYPE_BYTE = 0x00000000;
        public const uint PIPE_TYPE_MESSAGE = 0x00000004;
        public const uint PIPE_CLIENT_END = 0x00000000;
        public const uint PIPE_SERVER_END = 0x00000001;
        public const uint PIPE_UNLIMITED_INSTANCES = 255;
        public const uint NMPWAIT_WAIT_FOREVER = 0xffffffff;
        public const uint NMPWAIT_NOWAIT = 0x00000001;
        public const uint NMPWAIT_USE_DEFAULT_WAIT = 0x00000000;
        public const uint GENERIC_READ = (0x80000000);
        public const uint GENERIC_WRITE = (0x40000000);
        public const uint GENERIC_EXECUTE = (0x20000000);
        public const uint GENERIC_ALL = (0x10000000);
        public const uint CREATE_NEW = 1;
        public const uint CREATE_ALWAYS = 2;
        public const uint OPEN_EXISTING = 3;
        public const uint OPEN_ALWAYS = 4;
        public const uint TRUNCATE_EXISTING = 5;
        public const int INVALID_HANDLE_VALUE = -1;
        public const ulong ERROR_SUCCESS = 0;
        public const ulong ERROR_CANNOT_CONNECT_TO_PIPE = 2;
        public const ulong ERROR_PIPE_BUSY = 231;
        public const ulong ERROR_NO_DATA = 232;
        public const ulong ERROR_PIPE_NOT_CONNECTED = 233;
        public const ulong ERROR_MORE_DATA = 234;
        public const ulong ERROR_PIPE_CONNECTED = 535;
        public const ulong ERROR_PIPE_LISTENING = 536;

        [DllImport("kernel32.dll")]
        public static extern IntPtr CreateNamedPipe(
            String lpName,									// pipe name
            uint dwOpenMode,								// pipe open mode
            uint dwPipeMode,								// pipe-specific modes
            uint nMaxInstances,							// maximum number of instances
            uint nOutBufferSize,						// output buffer size
            uint nInBufferSize,							// input buffer size
            uint nDefaultTimeOut,						// time-out interval
            IntPtr pipeSecurityDescriptor		// SD
            );
        [DllImport("kernel32.dll")]
        public static extern bool ConnectNamedPipe(
            IntPtr hHandle,									// handle to named pipe
            Overlapped lpOverlapped					// overlapped structure
            );
        [DllImport("kernel32.dll")]
        public static extern bool CallNamedPipe(
            string lpNamedPipeName,
            byte[] lpInBuffer,
            uint nInBufferSize,
            byte[] lpOutBuffer,
            uint nOutBufferSize,
            byte[] lpBytesRead,
            int nTimeOut
            );
        [DllImport("kernel32.dll")]
        public static extern IntPtr CreateFile(
            String lpFileName,						  // file name
            uint dwDesiredAccess,					  // access mode
            uint dwShareMode,								// share mode
            SecurityAttributes attr,				// SD
            uint dwCreationDisposition,			// how to create
            uint dwFlagsAndAttributes,			// file attributes
            uint hTemplateFile);					  // handle to template file
        [DllImport("kernel32.dll")]
        public static extern bool ReadFile(
            IntPtr hHandle,											// handle to file
            byte[] lpBuffer,								// data buffer
            uint nNumberOfBytesToRead,			// number of bytes to read
            byte[] lpNumberOfBytesRead,			// number of bytes read
            uint lpOverlapped								// overlapped buffer
            );
        [DllImport("kernel32.dll")]
        public static extern bool WriteFile(
            IntPtr hHandle,											// handle to file
            byte[] lpBuffer,							  // data buffer
            uint nNumberOfBytesToWrite,			// number of bytes to write
            byte[] lpNumberOfBytesWritten,	// number of bytes written
            uint lpOverlapped								// overlapped buffer
            );
        [DllImport("kernel32.dll")]
        public static extern bool GetNamedPipeHandleState(
            IntPtr hHandle,
            IntPtr lpState,
            ref uint lpCurInstances,
            IntPtr lpMaxCollectionCount,
            IntPtr lpCollectDataTimeout,
            IntPtr lpUserName,
            IntPtr nMaxUserNameSize
            );
        [DllImport("kernel32.dll")]
        public static extern bool CancelIo(
            IntPtr hHandle
            );
        [DllImport("kernel32.dll")]
        public static extern bool WaitNamedPipe(
            String name,
            int timeout);
        [DllImport("kernel32.dll")]
        public static extern uint GetLastError();
        [DllImport("kernel32.dll")]
        public static extern bool FlushFileBuffers(
            IntPtr hHandle);
        [DllImport("kernel32.dll")]
        public static extern bool DisconnectNamedPipe(
            IntPtr hHandle);
        [DllImport("kernel32.dll")]
        public static extern bool SetNamedPipeHandleState(
            IntPtr hHandle,
            ref uint mode,
            IntPtr cc,
            IntPtr cd);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(
            IntPtr hHandle);
        private NamedPipeNative() { }
    }
    [StructLayout(LayoutKind.Sequential)]
    internal class SecurityAttributes
    {
    }
    [StructLayout(LayoutKind.Sequential)]
    internal class Overlapped
    {
    }
}
