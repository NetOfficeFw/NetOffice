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
    internal class TranslationClient
    {
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
            //Create an IPC client channel.
            IpcClientChannel channel = new IpcClientChannel();

            //Register the channel with ChannelServices.
            ChannelServices.RegisterChannel(channel, true);

            //Register the client type.
            RemotingConfiguration.RegisterWellKnownClientType(
                                typeof(WebTranslationService),
                                "ipc://NetOffice.SampleChannel/NetOffice.WebTranslationService.DataService");

            DataService = new WebTranslationService();
        }

        #endregion
    }
}
