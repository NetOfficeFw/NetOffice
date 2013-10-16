using System;
using System.Runtime.ConstrainedExecution;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace NetOffice.NamedPipes
{
    internal abstract class PipeConnection
    {
        protected PipeHandle Handle = new PipeHandle();
        protected string Name;
        protected bool disposed = false;
        protected int maxReadBytes;

        public string Read()
        {
            CheckIfDisposed();
            return NamedPipeWrapper.Read(Handle, maxReadBytes);
        }

        public byte[] ReadBytes()
        {
            CheckIfDisposed();
            return NamedPipeWrapper.ReadBytes(Handle, maxReadBytes);
        }

        public void Write(string text)
        {
            CheckIfDisposed();
            NamedPipeWrapper.Write(Handle, text);
        }

        public void WriteBytes(byte[] bytes)
        {
            CheckIfDisposed();
            NamedPipeWrapper.WriteBytes(Handle, bytes);
        }

        public abstract void Close();

        public abstract void Connect();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                NamedPipeWrapper.Close(this.Handle);
            }
            disposed = true;
        }

        public void CheckIfDisposed()
        {
            if (this.disposed)
            {
                throw new ObjectDisposedException("The Pipe Connection is disposed.");
            }
        }

        public InterProcessConnectionState GetState()
        {
            CheckIfDisposed();
            return this.Handle.State;
        }

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
