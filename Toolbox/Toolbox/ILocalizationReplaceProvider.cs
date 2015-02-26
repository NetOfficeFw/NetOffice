using System;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// Allows to replace embedded localization content
    /// </summary>
    public interface ILocalizationReplaceProvider
    {
        /// <summary>
        /// marker to replace content with
        /// </summary>
        /// <param name="marker">marker</param>
        /// <returns>replace conent or null</returns>
        string Replace(string marker);
    }
}
