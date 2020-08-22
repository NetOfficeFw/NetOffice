using System;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Tools;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// Connection part for the <see cref="ITaskPane"/> objects representing custom task panes.
    /// </summary>
    /// <typeparam name="T">Office Host Application</typeparam>
    public interface ITaskPaneConnection<in T> where T : ICOMObject
    {
        /// <summary>
        /// Called after startup to set the application instance and custom arguments to the custom task pane.
        /// </summary>
        /// <param name="application">host application instance</param>
        /// <param name="parentPane">custom task pane definition </param>
        /// <param name="customArguments">custom arguments</param>
        void OnConnection(T application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments);
    }

    /// <summary>
    /// Office TaskPane UserControl classes can implement this interface in NetOffice Tools Addin (COMAddin) as a special service.
    /// NetOffice will call ITaskPane members automatically.
    /// </summary>
    public interface ITaskPane : OfficeApi.Tools.ITaskPaneConnection<ICOMObject>
    {
        /// <summary>
        /// Called when Microsoft Office application is shuting down.
        /// The method is not called in case of unexpected termination of process.
        /// </summary>
        void OnDisconnection();

        /// <summary>
        /// Called when the user changes the dock position of the custom task pane.
        /// </summary>
        /// <param name="position">the current alignment for the instance</param>
        /// <remarks>
        /// This event is not fired when custom task pane size is changed.
        /// Use the <see cref="System.Windows.Forms.Control.Resize"/> event to listen for size changes.
        /// </remarks>
        void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position);

        /// <summary>
        /// Called when the user displays or closes the custom task pane.
        /// </summary>
        /// <param name="visible">the current visibility for the instance</param>
        /// <remarks>
        /// The <see cref="System.Windows.Forms.Control.VisibleChanged"/> event does not work as expected in a custom task pane scenarios.
        /// </remarks>
        void OnVisibleStateChanged(bool visible);
    }
}