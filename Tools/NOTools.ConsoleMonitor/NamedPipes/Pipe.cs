using System;
using System.Threading;
using System.IO;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    internal sealed class Pipe : IDisposable
    {
        #region Fields

        internal Thread PipeThread;
        internal ServerPipeConnection PipeConnection;
        internal bool Listen = true;
        internal DateTime LastAction;

        private bool _disposed = false;
        private IChannelManager _parent;

        #endregion

        #region Construction

        internal Pipe(IChannelManager parent, string name, uint outBuffer, uint inBuffer, int maxReadBytes)
        {
            _parent = parent;
            PipeConnection = new ServerPipeConnection(name, outBuffer, inBuffer, maxReadBytes);
            PipeThread = new Thread(new ThreadStart(PipeListener));
            PipeThread.IsBackground = true;
            PipeThread.Name = "Pipe Thread " + this.PipeConnection.NativeHandle.ToString();
            LastAction = DateTime.Now;
        }

        #endregion

        private void PipeListener()
        {
            CheckIfDisposed();
            try
            {
                Listen = _parent.Listen;
                while (Listen)
                {
                    LastAction = DateTime.Now;
                    string request = PipeConnection.Read();
                    LastAction = DateTime.Now;
                    if (request.Trim() != "")
                        PipeConnection.Write(_parent.HandleRequest(request));
                    else
                        PipeConnection.Write("");

                    LastAction = DateTime.Now;
                    PipeConnection.Disconnect();
                    if (Listen)
                        Connect();

                    _parent.WakeUp();
                }
            }
            catch (System.Threading.ThreadAbortException) { }
            catch (System.Threading.ThreadStateException) { }
            catch{}
            finally
            {
                this.Close();
            }
        }

        internal void Connect()
        {
            CheckIfDisposed();
            PipeConnection.Connect();
        }

        internal void Close()
        {
            CheckIfDisposed();
            this.Listen = false;
            _parent.RemoveServerChannel(this.PipeConnection.NativeHandle);
            this.Dispose();
        }

        internal void Start()
        {
            CheckIfDisposed();
            PipeThread.Start();
        }

        #region IDisposable Members

        private void CheckIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException("ServerNamedPipe");
        }

        ~Pipe()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                PipeConnection.Dispose();
                if (PipeThread != null)
                {
                    try
                    {
                        PipeThread.Abort();
                    }
                    catch (Exception ex) { Console.WriteLine(ex.Message); }
                }
            }
            _disposed = true;
        }

        #endregion
    }
}