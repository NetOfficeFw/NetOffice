using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.NamedPipes
{
    internal sealed class PipeHandle
    {
        public IntPtr Handle;
        public InterProcessConnectionState State;

        public PipeHandle(int hnd)
        {
            this.Handle = new IntPtr(hnd);
            this.State = InterProcessConnectionState.NotSet;
        }

        public PipeHandle(int hnd, InterProcessConnectionState state)
        {
            this.Handle = new IntPtr(hnd);
            this.State = state;
        }

        public PipeHandle()
        {
            this.Handle = new IntPtr(NamedPipeNative.INVALID_HANDLE_VALUE);
            this.State = InterProcessConnectionState.NotSet;
        }
    }
}
