using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NetOffice.NamedPipes
{
    internal sealed class ClientPipeConnection : PipeConnection
    {
        private string Server = ".";

        public override void Close()
        {
            CheckIfDisposed();
            NamedPipeWrapper.Close(this.Handle);
        }

        public override void Connect()
        {
            CheckIfDisposed();
            this.Handle = NamedPipeWrapper.ConnectToPipe(this.Name, this.Server);
        }

        public bool TryConnect()
        {
            CheckIfDisposed();
            bool ReturnVal = NamedPipeWrapper.TryConnectToPipe(this.Name, this.Server, out this.Handle);
            return ReturnVal;
        }

        public ClientPipeConnection(string name)
        {
            this.Name = name;
            this.Server = ".";
            this.maxReadBytes = Int32.MaxValue;
        }

        public ClientPipeConnection(string name, string server)
        {
            this.Name = name;
            this.Server = server;
            this.maxReadBytes = Int32.MaxValue;
        }

        ~ClientPipeConnection()
        {
            Dispose(false);
        }
    }
}
