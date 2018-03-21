using System;

namespace NetOffice
{
    /// <summary>
    /// Implementing instance supports automatic quit in dispose.
    /// </summary>
    public interface IAutomaticQuit
    {
        /// <summary>
        /// Determines Quit method want be called while disposing if NetOffice.Settings.EnableAutomaticQuit is true.
        /// Default is true when instance has no parent object and its not a cloned instance, otherwise false.
        /// </summary>
        bool Enabled { get; set; }
    }
}