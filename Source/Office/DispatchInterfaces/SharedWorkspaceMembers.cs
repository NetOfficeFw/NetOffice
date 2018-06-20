using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceMembers 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861050.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0382-0000-0000-C000-000000000046")]
    public interface SharedWorkspaceMembers : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceMember>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.SharedWorkspaceMember this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861505.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863728.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863376.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool ItemCountExceeded { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860533.aspx </remarks>
        /// <param name="email">string email</param>
        /// <param name="domainName">string domainName</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="role">optional object role</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceMember Add(string email, string domainName, string displayName, object role);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860533.aspx </remarks>
        /// <param name="email">string email</param>
        /// <param name="domainName">string domainName</param>
        /// <param name="displayName">string displayName</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.SharedWorkspaceMember Add(string email, string domainName, string displayName);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SharedWorkspaceMember>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.SharedWorkspaceMember> GetEnumerator();

        #endregion
    }
}
