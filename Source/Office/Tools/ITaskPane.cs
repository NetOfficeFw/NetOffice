using System;
using NetOffice;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// UserControls for a CustomTaskPane can implement these interface. The COMAddin class call the methods.
    /// </summary>
    public interface ITaskPane
    {
        /// <summary>
        /// Called from the COMAddin class while creation in CTPFactoryAvailable
        /// </summary>
        /// <param name="application">Host Application Instance</param>
		/// <param name="customArguments">optional arguments</param>
        void OnConnection(COMObject application, object[] customArguments);

		/// <summary>
        /// Called from the COMAddin class while OnDisconnection
        /// </summary>
		void OnDisconnection();
    }
}