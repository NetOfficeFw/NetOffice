using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.NamedPipes
{
    internal class PipeClient
    {
        private static string _pipeName = "NOTools.ConsoleMonitor.PipeConnection";

        private static int _maxMessageLenght = 1024;

        /// <summary>
        /// Send a message to specific console
        /// </summary>
        /// <param name="console">name of the console. use null for main console</param>
        /// <param name="message">given message as any</param>        
        /// <returns>loghandle from server if recieved</returns>
        public string SendConsoleMessage(string console, string message)
        {
            return SendConsoleMessage(console, message, "");
        }

        /// <summary>
        /// Send a message to specific console
        /// </summary>
        /// <param name="console">name of the console. use null for main console</param>
        /// <param name="message">given message as any</param>  
        /// <param name="parentMessageID">parent loghandle or null</param>  
        /// <returns>loghandle from server if recieved</returns>
        public string SendConsoleMessage(string console, string message, string parentMessageID)
        {
            if (null != console && console.IndexOf("?") < -1)
                throw new ArgumentException("console must be without '?' character");

            if (String.IsNullOrEmpty(message) || message.Length > _maxMessageLenght)
                return null;

            if (null == parentMessageID)
                parentMessageID = "";

            DateTime now = DateTime.Now;
            string timeString = now.ToLongTimeString() + ":" + now.Millisecond;

            return SendRecieveString("CNSL?" + console + "?" + Environment.MachineName + "?" + (null != AppDomain.CurrentDomain ? AppDomain.CurrentDomain.FriendlyName + AppDomain.CurrentDomain.Id.ToString() : "") + "?" + timeString + "?" + parentMessageID + "?" + message);
        }


        /// <summary>
        /// Send a message to specific channel
        /// </summary>
        /// <param name="channel">name of the channel</param>
        /// <param name="message">given message as any</param>
        /// <returns>loghandle from server if recieved</returns>
        public string SendChannelMessage(string channel, string message)
        {
            if (String.IsNullOrEmpty(channel) || channel.IndexOf("?") < -1)
                throw new ArgumentException("channel can't empty und must be without '?' character");
            if (String.IsNullOrEmpty(message) || message.Length > _maxMessageLenght)
                return null;

            DateTime now = DateTime.Now;
            string timeString = now.ToLongTimeString() + ":" + now.Millisecond;
            return SendRecieveString("CHNL?" + channel + "?" + Environment.MachineName + "?" + (null != AppDomain.CurrentDomain ? AppDomain.CurrentDomain.FriendlyName + AppDomain.CurrentDomain.Id.ToString() : "") + "?" + timeString + "?" + /*parentMessageID*/ "?" + message);
        }


        private string SendRecieveString(string any)
        {
            ClientPipeConnection clientConnection = null;
            try
            {
                clientConnection = new ClientPipeConnection(_pipeName, ".");
                if (!clientConnection.TryConnect())
                {
                    clientConnection.Dispose();
                    return null;
                }

                clientConnection.Write(any);
                string response = clientConnection.Read();
                clientConnection.Close();
                clientConnection.Dispose();
                return response;
            }
            catch (Exception exception)
            {
                if (null != clientConnection)
                    clientConnection.Dispose();
                throw (exception);
            }
        }
    }
}
