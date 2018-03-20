using System;

namespace NetOffice
{
    /// <summary>
    /// Represents an application version
    /// </summary>
    public interface IApplicationVersionProvider
    {
        /// <summary>
        /// Unique name of the hosting component. For example: 'NetOffice.ExcelApi' or 'NetOffice.WordApi'
        /// </summary>
        string ComponentName { get; }

        /// <summary>
        /// The display name of the implementation that provides the version
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Request version information on demand and cache to call the remote server only 1x times
        /// </summary>
        object Version { get; }

        /// <summary>
        /// Determines version is already requested
        /// </summary>
        bool VersionRequested { get; }

        /// <summary>
        /// Force try update the version even version is already requested
        /// </summary>
        void TryRequestVersion();
    }
}