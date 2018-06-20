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
    /// DispatchInterface _CustomXMLSchemaCollection 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000CDB02-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OfficeApi.CustomXMLSchemaCollection))]
    public interface _CustomXMLSchemaCollection : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CustomXMLSchema>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862208.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864015.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.CustomXMLSchema this[object index] { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860878.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_NamespaceURI(Int32 index);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_NamespaceURI
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860878.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_NamespaceURI")]
        string NamespaceURI(Int32 index);

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="alias">optional string Alias = </param>
        /// <param name="fileName">optional string FileName = </param>
        /// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias, object fileName, object installForAllUsers);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLSchema Add();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="alias">optional string Alias = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864881.aspx </remarks>
        /// <param name="namespaceURI">optional string NamespaceURI = </param>
        /// <param name="alias">optional string Alias = </param>
        /// <param name="fileName">optional string FileName = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLSchema Add(object namespaceURI, object alias, object fileName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864690.aspx </remarks>
        /// <param name="schemaCollection">NetOffice.OfficeApi.CustomXMLSchemaCollection schemaCollection</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void AddCollection(NetOffice.OfficeApi.CustomXMLSchemaCollection schemaCollection);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864142.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        bool Validate();

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CustomXMLSchema>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.CustomXMLSchema> GetEnumerator();

        #endregion
    }
}
