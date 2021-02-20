using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// Represents information about a disabled add-in in a Microsoft Office application.
    /// </summary>
    /// <remarks>
    /// Use the <see cref="OfficeResiliency.Parse()"/> method to convert the binary
    /// data from the Resiliency\DisabledItems registry keys into the <see cref="DisabledItem"/> object.
    /// </remarks>
    public class DisabledItem
    {
        public DisabledItemType DisabledItemType { get; set; }

        public string FriendlyName { get; set; }

        public string Module { get; set; }
    }
}