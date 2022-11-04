using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// Represents information about a disabled add-in in a Microsoft Office application.
    /// </summary>
    /// <remarks>
    /// Use the <see cref="OfficeResiliency.Parse(byte[])"/> method to convert the binary
    /// data from the Resiliency\DisabledItems registry keys into the <see cref="DisabledItem"/> object.
    /// </remarks>
    public class DisabledItem
    {
        /// <summary>
        /// Type of the disabled item.
        /// </summary>
        public DisabledItemType DisabledItemType { get; set; }

        /// <summary>
        /// Friendly name of the disabled add-in.
        /// </summary>
        public string FriendlyName { get; set; }

        /// <summary>
        /// Module name of the disabled add-in.
        /// </summary>
        public string Module { get; set; }
    }
}