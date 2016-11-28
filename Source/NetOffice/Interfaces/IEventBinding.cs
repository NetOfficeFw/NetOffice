using System;
using System.ComponentModel;
namespace NetOffice
{
    /// <summary>
    /// EventBinding Interface
    /// </summary>
    public interface IEventBinding
    {
        /// <summary>
        /// Event bridge is advised
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool EventBridgeInitialized { get; }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool HasEventRecipients();

        /// <summary>
        /// Recipient delegates for an event
        /// </summary>
        /// <param name="eventName">name of the even</param>
        /// <returns>recipients delegates</returns>
        Delegate[] GetEventRecipients(string eventName);

        /// <summary>
        /// Instance has one or more event recipients for a specific event
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns>the count of recipients</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        int GetCountOfEventRecipients(string eventName);

        /// <summary>
        /// Call a specific event for all recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <param name="paramsArray">argument array</param>
        /// <returns>count of recipients</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        int RaiseCustomEvent(string eventName, ref object[] paramsArray);

        /// <summary>
        /// Create the event eventbridge for the object
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void CreateEventBridge();

        /// <summary>
        /// Dispose the event eventbridge for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void DisposeEventBridge();
    }
}
