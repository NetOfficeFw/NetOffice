using System;
using Publisher = NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Tools
{
    /// <summary>
    /// Task pane UserControl instances can implement these interface in a NetOffice Tools Addin as a special service
    /// </summary>
    public interface ITaskPane : OfficeApi.Tools.ITaskPaneConnection<Publisher.Application>
    {
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
