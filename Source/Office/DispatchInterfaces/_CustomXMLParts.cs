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
    /// DispatchInterface _CustomXMLParts 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
    public interface _CustomXMLParts : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CustomXMLPart>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862384.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865208.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.CustomXMLPart this[object index] { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865361.aspx </remarks>
        /// <param name="xML">optional string XML = </param>
        /// <param name="schemaCollection">optional object schemaCollection</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLPart Add(object xML, object schemaCollection);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865361.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLPart Add();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865361.aspx </remarks>
        /// <param name="xML">optional string XML = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLPart Add(object xML);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865246.aspx </remarks>
        /// <param name="id">string id</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLPart SelectByID(string id);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861183.aspx </remarks>
        /// <param name="namespaceURI">string namespaceURI</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        NetOffice.OfficeApi.CustomXMLParts SelectByNamespace(string namespaceURI);

        #endregion
    }
}
