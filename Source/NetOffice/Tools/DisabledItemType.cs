using System;

namespace NetOffice.Tools
{
    /// <summary>
    /// Represents the type of structure stored in the Resiliency registry keys.
    /// </summary>
    /// <remarks>
    /// The Resiliency registry key contains a list of DisabledItems. Information
    /// about disabled items (eg. add-ins) is stored as binary data.
    /// The first <see cref="Int32"/> value represents the type of the
    /// data structure.
    /// 
    /// For regular add-ins (COM, .NET, NetOffice or VSTO) the usual disabled item
    /// type value would be either AddInByFilename (1) or AddInByDEPFilename (6).
    /// </remarks>
    public enum DisabledItemType
    {
        None = 0,
        AddInByFilename = 1,
        DocumentByPath = 2,
        OpenFailedFilename = 2,
        Workpane = 3,
        COMObject = 4,
        COMDEPObject = 5,
        AddInByDEPFilename = 6,
        Printer = 7,
        PrintTicket = 8,
        AppSpecificItems = 0x40000000,
        OutlookDisablePreviewPane = 0x40000000,
        OutlookAddinAutoDisabledByFileName = 1073741825
    }
}
