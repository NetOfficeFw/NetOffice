using System;
using Word = NetOffice.WordApi;

namespace NetOffice.WordApi.Tools
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
        void OnConnection(Word.Application application, object[] customArguments);
    }
}