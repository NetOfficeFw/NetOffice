using System;
using Publisher = NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Tools
{
    /// <summary>
    /// Custom task pane UserControl instance may implement this interface to be notified about the lifetime of the custom task pane.
    /// </summary>
    public interface ITaskPane : OfficeApi.Tools.ITaskPaneConnection<Publisher.Application>
    {
        /// <summary>
        /// Called when Microsoft Office application is shuting down. This method is not called in case of unexpected termination of the process.
        /// </summary>
        void OnDisconnection();

        /// <summary>
        /// Called when the user changes the dock position of the custom task pane.
        /// </summary>
        /// <param name="position">the current alignment for the instance</param>
        void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position);

        /// <summary>
        /// Called when the user displays or closes the custom task pane.
        /// </summary>
        /// <param name="visible">the current visibility for the instance</param>
        void OnVisibleStateChanged(bool visible);
    }
}
