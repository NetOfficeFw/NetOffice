using System;

namespace NetOffice.OutlookApi.Tools
{
    /// <summary>
    /// Indicates in which method the error is occured
    /// </summary>
    public enum OutlookErrorMethodKind
    {
        /// <summary>
        /// the error is occured in OpenForm_Close
        /// </summary>
        CloseOpenFormRegion = 0,

        /// <summary>
        /// the error is occured in GetFormRegionIcon
        /// </summary>
        GetFormRegionIcon = 1,

        /// <summary>
        /// the error is occured in GetFormRegionManifest
        /// </summary>
        GetFormRegionManifest = 2,

        /// <summary>
        /// the error is occured in BeforeFormRegionShow
        /// </summary>
        BeforeFormRegionShow = 3,

        /// <summary>
        /// the error is occured in GetFormRegionStorage
        /// </summary>
        GetFormRegionStorage = 4
    }
}
