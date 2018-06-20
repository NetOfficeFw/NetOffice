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
    /// DispatchInterface Permission 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861518.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C0376-0000-0000-C000-000000000046")]
    public interface Permission : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.UserPermission>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.UserPermission this[object index] { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865565.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862116.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool EnableTrustedBrowser { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861383.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865228.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool Enabled { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862755.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string RequestPermissionURL { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860910.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string PolicyName { get; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860601.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string PolicyDescription { get;}

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864131.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool StoreLicenses { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864905.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        string DocumentAuthor { get; set; }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863690.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        bool PermissionFromPolicy { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863139.aspx </remarks>
        /// <param name="userId">string userId</param>
        /// <param name="permission">optional object permission</param>
        /// <param name="expirationDate">optional object expirationDate</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.UserPermission Add(string userId, object permission, object expirationDate);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863139.aspx </remarks>
        /// <param name="userId">string userId</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.UserPermission Add(string userId);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863139.aspx </remarks>
        /// <param name="userId">string userId</param>
        /// <param name="permission">optional object permission</param>
        [CustomMethod]
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.UserPermission Add(string userId, object permission);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864678.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void ApplyPolicy(string fileName);

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861135.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        void RemoveAll();

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.UserPermission>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.UserPermission> GetEnumerator();

        #endregion
    }
}
