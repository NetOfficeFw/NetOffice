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
		/// <param name="parentPane">CustomTaskPane definition </param>
		/// <param name="customArguments">optional arguments</param>
        void OnConnection(COMObject application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments);

		/// <summary>
        /// Called from the COMAddin class while OnDisconnection
        /// </summary>
		void OnDisconnection();
    }
}