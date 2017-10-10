using System;

namespace NetOffice.OutlookApi.Tools.Contribution.Security
{
    /// <summary>
    /// Click strategy for the security dialog suppressor
    /// </summary>
    public enum ClickStrategy
    {
        /// <summary>
        /// Move mouse to position, perform click and restore origin mouse position
        /// </summary>
        MoveTo = 0,

        /// <summary>
        /// Use SendMessage
        /// </summary>
        SendMessage = 1,

        /// <summary>
        /// Use PostMessage
        /// </summary>
        PostMessage = 2,

        /// <summary>
        /// Do nothing, Use Suppress.OnAction event to handle click at hand
        /// </summary>
        None = 3
    }
}
