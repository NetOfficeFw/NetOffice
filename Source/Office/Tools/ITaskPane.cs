using System;
using NetOffice;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// ITaskPane Connection Part
    /// </summary>
    /// <typeparam name="T">Office Host Application</typeparam>
    public interface ITaskPaneConnection<T> where T : ICOMObject
    {
        /// <summary>
        /// After startup to serve the application instance and custom arguments(if set)
        /// </summary>
        /// <param name="application">host application instance</param>
		/// <param name="parentPane">custom task pane definition </param>
		/// <param name="customArguments">custom arguments</param>
        void OnConnection(T application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments);
    }

    /// <summary>
    /// Office TaskPane UserControl classes can implement these interface in a NetOffice Tools Addin(COMAddin) as a special service.
    /// NetOffice want call ITaskPane members automaticly 
    /// </summary>
    public interface ITaskPane : OfficeApi.Tools.ITaskPaneConnection<ICOMObject>
    {
		/// <summary>
        /// While Office Application shutdown. The method is not called in case of unexpected termination (may user kills the instance in task manager)
        /// </summary>
		void OnDisconnection();

        /// <summary>
        /// Called after any position changes but not for size changes. Use the UserControl.Resize event instead for size changes
        /// </summary>
        /// <param name="position">the current alignment for the instance</param>
        void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position);

        /// <summary>
        /// Called after any visibility changes because the UserControl.VisibleChanged event doesnt work as expected in a task pane scenario
        /// </summary>
        /// <param name="visible">the current visibility for the instance</param>
        void OnVisibleStateChanged(bool visible);
    }
}