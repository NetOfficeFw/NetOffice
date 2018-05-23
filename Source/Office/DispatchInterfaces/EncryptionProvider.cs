using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface EncryptionProvider 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863389.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface EncryptionProvider : ICOMObject
    {
        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864896.aspx </remarks>
        /// <param name="encprovdet">NetOffice.OfficeApi.Enums.EncryptionProviderDetail encprovdet</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        object GetProviderDetail(NetOffice.OfficeApi.Enums.EncryptionProviderDetail encprovdet);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864027.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 NewSession(object parentWindow);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864627.aspx </remarks>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="encryptionData">object encryptionData</param>
        /// <param name="permissionsMask">UIntPtr permissionsMask</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Authenticate(object parentWindow, object encryptionData, out UIntPtr permissionsMask);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864902.aspx </remarks>
        /// <param name="sessionHandle">Int32 sessionHandle</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 CloneSession(Int32 sessionHandle);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864652.aspx </remarks>
        /// <param name="sessionHandle">Int32 sessionHandle</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void EndSession(Int32 sessionHandle);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862766.aspx </remarks>
        /// <param name="sessionHandle">Int32 sessionHandle</param>
        /// <param name="encryptionData">object encryptionData</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Save(Int32 sessionHandle, object encryptionData);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861839.aspx </remarks>
        /// <param name="sessionHandle">Int32 sessionHandle</param>
        /// <param name="streamName">string streamName</param>
        /// <param name="unencryptedStream">object unencryptedStream</param>
        /// <param name="encryptedStream">object encryptedStream</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void EncryptStream(Int32 sessionHandle, string streamName, object unencryptedStream, object encryptedStream);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864940.aspx </remarks>
        /// <param name="sessionHandle">Int32 sessionHandle</param>
        /// <param name="streamName">string streamName</param>
        /// <param name="encryptedStream">object encryptedStream</param>
        /// <param name="unencryptedStream">object unencryptedStream</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void DecryptStream(Int32 sessionHandle, string streamName, object encryptedStream, object unencryptedStream);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863378.aspx </remarks>
        /// <param name="sessionHandle">Int32 sessionHandle</param>
        /// <param name="parentWindow">object parentWindow</param>
        /// <param name="readOnly">bool readOnly</param>
        /// <param name="remove">bool remove</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void ShowSettings(Int32 sessionHandle, object parentWindow, bool readOnly, out bool remove);

        #endregion
    }
}
