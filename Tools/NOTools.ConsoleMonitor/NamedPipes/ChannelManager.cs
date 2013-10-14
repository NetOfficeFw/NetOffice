using System;
using System.Collections;
using System.Threading;
using System.Web;
using System.IO;
using System.Configuration;
using System.Diagnostics;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    public delegate void RequestEventHandler(string request, ref string response);

    public class ChannelManager : IChannelManager
    {
        public Hashtable Pipes;
        public object SyncRoot = new object();

        private uint NumberPipes = 5;
        private uint OutBuffer = 512;
        private uint InBuffer = 512;

        private const int MAX_READ_BYTES = 5000;
        private const int PIPE_MAX_STUFFED_TIME = 1000;

        private bool _listen = true;
        private int _numChannels = 0;
        private Hashtable _pipes = new Hashtable();
        private Thread _mainThread;
        private string _pipeName;
        private ManualResetEvent _mre;

        public event RequestEventHandler Request;

        internal ChannelManager(string pipeName)
        {
            _pipeName = pipeName;
        }

        public bool Listen
        {
            get
            {
                return _listen;
            }
            set
            {
                _listen = value;
            }
        }

        public void Initialize()
        {
            Pipes = Hashtable.Synchronized(_pipes);
            _mre = new ManualResetEvent(false);
            _mainThread = new Thread(new ThreadStart(Start));
            _mainThread.IsBackground = true;
            _mainThread.Name = "Main Pipe Thread";
            _mainThread.Start();
            Thread.Sleep(1000);
        }

        public string HandleRequest(string request)
        {
            if (null != Request)
            {
                string response = "";
                Request(request, ref response);
                if (null == response)
                    response = "";
                return response;
            }
            else
                return "";
        }

        private void Start()
        {
            try
            {
                while (_listen)
                {
                    int[] keys = new int[Pipes.Keys.Count];
                    Pipes.Keys.CopyTo(keys, 0);
                    foreach (int key in keys)
                    {
                        Pipe serverPipe = (Pipe)Pipes[key];
                        if (serverPipe != null && DateTime.Now.Subtract(serverPipe.LastAction).Milliseconds > 
                            PIPE_MAX_STUFFED_TIME && serverPipe.PipeConnection.GetState() != InterProcessConnectionState.WaitingForClient)
                        {
                            serverPipe.Listen = false;
                            serverPipe.PipeThread.Abort();
                            RemoveServerChannel(serverPipe.PipeConnection.NativeHandle);
                        }
                    }

                    if (_numChannels <= NumberPipes)
                    {
                        Pipe pipe = new Pipe(this, _pipeName, OutBuffer, InBuffer, MAX_READ_BYTES);
                        try
                        {
                            pipe.Connect();
                            pipe.LastAction = DateTime.Now;
                            System.Threading.Interlocked.Increment(ref _numChannels);
                            pipe.Start();
                            if(!pipe.PipeConnection.IsDisposed)
                                Pipes.Add(pipe.PipeConnection.NativeHandle, pipe);
                        }
                        catch (InterProcessIOException)
                        {
                            RemoveServerChannel(pipe.PipeConnection.NativeHandle);
                            pipe.Dispose();
                        }
                    }
                    else
                    {
                        _mre.Reset();
                        _mre.WaitOne(1000, false);
                    }
                }
            }
            catch
            {
                // Log exception
            }
        }

        public void Stop()
        {
            _listen = false;
            _mre.Set();
            try
            {
                int[] keys = new int[Pipes.Keys.Count];
                Pipes.Keys.CopyTo(keys, 0);
                foreach (int key in keys)
                {
                    ((Pipe)Pipes[key]).Listen = false;
                }
                int i = _numChannels * 3;
                for (int j = 0; j < i; j++)
                {
                    StopServerPipe();
                }
                Pipes.Clear();
                _mre.Close();
                _mre = null;
            }
            catch
            {
                // Log exception
            }
        }

        public void WakeUp()
        {
            if (_mre != null)
            {
                _mre.Set();
            }
        }

        private void StopServerPipe()
        {
            try
            {
                ClientPipeConnection pipe = new ClientPipeConnection(_pipeName);
                if (pipe.TryConnect())
                {
                    pipe.Close();
                }
            }
            catch
            {
                // Log exception
            }
        }

        public void RemoveServerChannel(object param)
        {
            int handle = (int)param;
            System.Threading.Interlocked.Decrement(ref _numChannels);
            Pipes.Remove(handle);
            this.WakeUp();
        }
    }
}