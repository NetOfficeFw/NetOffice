using System;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;
using Sample.Server;

namespace Sample.ExcelAddin
{
    /// <summary>
    /// These class handles the IPC communication and contains the WebTranslationService proxy
    /// </summary>
    internal class TranslationClient : IDisposable
    {
        #region Fields

        private IpcClientChannel _channel;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal TranslationClient()
        {
            RegisterProxy();
        }

        #endregion

        #region Properties

        /// <summary>
        /// IPC Client Proxy to the Translation Service
        /// </summary>
        internal WebTranslationService DataService { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// Get information the server is available
        /// </summary>
        /// <returns>true if available</returns>
        internal bool IsValid()
        {
            try
            {
                var dumy = DataService.AvailableTranslations;
                return true;
            }
            catch (Exception exception)
            {
                Console.Write(string.Format("A {0} occured in TranslationClient.IsValid", exception.GetType().Name));
                return false;
            }
        }   
         
        /// <summary>
        /// Register the ipc client proxy
        /// </summary>
        internal void RegisterProxy()
        {
            try
            {
                string uri = "ipc://NetOffice.SampleChannel/NetOffice.WebTranslationService.DataService";

                //Create an IPC client channel.
                _channel = new IpcClientChannel();

                //Register the channel with ChannelServices.
                ChannelServices.RegisterChannel(_channel, true);

                //Register the client type.
                WellKnownClientTypeEntry[] entries = RemotingConfiguration.GetRegisteredWellKnownClientTypes();
                if (null == GetEntry(entries, uri))
                {
                    RemotingConfiguration.RegisterWellKnownClientType(
                                        typeof(WebTranslationService),
                                        uri);
                }
                DataService = new WebTranslationService();

                // try to do some action to see the server is alive
                string[] dumy = DataService.AvailableTranslations;
            }
            catch (RemotingException exception)
            {
                // rethrow the exception with a friendly message
                throw new RemotingException("Unable to connect the local translation service.", exception);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Get entry from array based on objectUrl or null if not match. We can not use Linq here because NO examples comes also in .Net 2
        /// </summary>
        /// <param name="entries">entries from RemotingConfiguration</param>
        /// <param name="objectUrl">target url</param>
        /// <returns>entry or null</returns>
        private static WellKnownClientTypeEntry GetEntry(WellKnownClientTypeEntry[] entries, string objectUrl)
        {
            foreach (WellKnownClientTypeEntry item in entries)
            {
                if (objectUrl.Equals(item.ObjectUrl, StringComparison.InvariantCultureIgnoreCase))
                    return item;
            }
            return null;
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (null != _channel)
            {
                ChannelServices.UnregisterChannel(_channel);
                _channel = null;
            }
        }

        #endregion
    }
}
