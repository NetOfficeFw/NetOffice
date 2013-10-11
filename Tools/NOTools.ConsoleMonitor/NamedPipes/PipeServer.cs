using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Security.Principal;
using System.Diagnostics;

namespace NOTools.ConsoleMonitor.NamedPipes
{
    internal class PipeServer : INotificationProvider
    {
        #region Fields

        protected internal ChannelManager _channelManager;

        #endregion

        #region Construction

        public PipeServer(string pipeName = "NOTools.ConsoleMonitor.PipeConnection")
        {
            _channelManager = new ChannelManager(pipeName);
            _channelManager.Request += new RequestEventHandler(ChannelManager_Request);
            Start();
        }
        
        #endregion

        #region Properties

        public bool IsRunning { get; private set; }

        #endregion

        #region Methods

        public void Start()
        {
            if (!IsRunning)
            {
                _channelManager.Initialize();
                IsRunning = true;
            }
        }

        public void Stop()
        {
            if (IsRunning)
            {
                _channelManager.Stop();
                IsRunning = false;
            }
        }
        
        #endregion

        private static string[] CreateMagicArray(string request)
        { 
            string[] array = new string[6];

            string[] splitArray = request.Split(new string[] { "?" }, StringSplitOptions.None);
            if (splitArray.Length < 7)
                return null;
            array[0] = splitArray[1]; // console/channel name            
            array[1] = splitArray[2]; // machine name
            array[2] = splitArray[3]; // appdomain name
            array[3] = splitArray[4]; // timecode
            array[4] = splitArray[5]; // parent message id
            array[5] = splitArray[6]; // given message
            return array;
        }

        #region Trigger

        private void ChannelManager_Request(string request, ref string response)
        {
            if (!IsRunning)
                return;
            if (!String.IsNullOrWhiteSpace(request))
            {
                string[] magicArray = CreateMagicArray(request);
                if (null == magicArray)
                    return;

                if (request.StartsWith("CNSL"))
                    response = RaiseConsoleNotification(magicArray[3], magicArray[0], magicArray[1], magicArray[2], magicArray[4], magicArray[5]);
                else if (request.StartsWith("CHNL"))
                    response = RaiseChannelNotification(magicArray[3], magicArray[0], magicArray[1], magicArray[2], magicArray[4], magicArray[5]);
            }
        }

        #endregion

        #region INotification

        public event UpdateMonitorInvoker ConsoleNotification;

        private string RaiseConsoleNotification(string notifyTime, string consoleName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime = false, bool showMachine = false, bool showAppDomain = false)
        {
            string newEntryID = null;
            if (null != ConsoleNotification)
                newEntryID = ConsoleNotification(notifyTime, consoleName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);
            return newEntryID;
        }

        public event UpdateMonitorInvoker ChannelNotification;

        private string RaiseChannelNotification(string notifyTime, string channelName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime = false, bool showMachine = false, bool showAppDomain = false)
        {
            string newEntryID = null;
            if (null != ChannelNotification)
                newEntryID = ChannelNotification(notifyTime, channelName, machineName, appDomainFriendlyName, parentEntryID, message, showTime, showMachine, showAppDomain);
            return newEntryID;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (null != _channelManager && _channelManager.Listen)
            {
                _channelManager.Stop();
                _channelManager = null;
            }
        }

        #endregion
    }
}
