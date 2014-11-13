using System;
using NetOffice;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools
{
 	/// <summary>
    /// Task pane UserControl instances can implement these interface in a NetOffice Tools Addin as a special service
    /// </summary>
    public interface ITaskPane
    {
        /// <summary>
        /// After startup to serve the application instance and custom arguments(if set)
        /// </summary>
        /// <param name="application">host application instance</param>
		/// <param name="parentPane">custom task pane definition </param>
		/// <param name="customArguments">custom arguments</param>
        void OnConnection(COMObject application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments);

		/// <summary>
        /// While Excel Application shutdown. The method is not called in case of unexpected termination (may user kill the instance in task manager)
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