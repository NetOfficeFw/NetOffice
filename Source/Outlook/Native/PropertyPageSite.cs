using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi.Native
{
    /// <summary>
    /// NativeInterface PropertyPageSite SupportByVersion Outlook, 9,10,11,12,14,15,16
    /// Represents the container of a custom property page.
    /// </summary>
    [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
    [ComImport, Guid("0006307F-0000-0000-C000-000000000046"), TypeLibType(4160)]
    [EntityType(EntityType.IsNativeInterface)]
    public interface PropertyPageSite
    {
        /// <summary>
        /// Returns an Application object that represents the parent Outlook application for the object. Read-only.
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [DispId(61440)]
        object Application
        {
            [DispId(61440)]
            [MethodImpl(4096)]
            [return: MarshalAs(28)]
            get;
        }

        /// <summary>
        /// Returns an OlObjectClass constant indicating the object's class. Read-only.
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [DispId(61450)]
        object Class
        {
            [DispId(61450)]
            [MethodImpl(4096)]
            get;
        }

        /// <summary>
        /// Returns the NameSpace object for the current session. Read-only.
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [DispId(61451)]
        object Session
        {
            [DispId(61451)]
            [MethodImpl(4096)]
            [return: MarshalAs(28)]
            get;
        }

        /// <summary>
        /// Returns the parent Object of the specified object. Read-only.
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [DispId(61441)]
        object Parent
        {
            [DispId(61441)]
            [MethodImpl(4096)]
            [return: MarshalAs(26)]
            get;
        }

        /// <summary>
        /// Notifies Microsoft Outlook that a custom property page has changed.
        /// </summary>
        [SupportByVersion("Outlook", 9, 10, 11, 12, 14, 15, 16)]
        [DispId(8448)]
        [MethodImpl(4096)]
        void OnStatusChange();
    }
}
