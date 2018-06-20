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
    /// DispatchInterface CustomXMLValidationErrors 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860565.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000CDB0F-0000-0000-C000-000000000046")]
    public interface CustomXMLValidationErrors : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CustomXMLValidationError>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862460.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861469.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Count { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.CustomXMLValidationError this[Int32 index] { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860763.aspx </remarks>
        /// <param name="node">NetOffice.OfficeApi.CustomXMLNode node</param>
        /// <param name="errorName">string errorName</param>
        /// <param name="errorText">optional string ErrorText = </param>
        /// <param name="clearedOnUpdate">optional bool ClearedOnUpdate = true</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Add(NetOffice.OfficeApi.CustomXMLNode node, string errorName, object errorText, object clearedOnUpdate);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860763.aspx </remarks>
        /// <param name="node">NetOffice.OfficeApi.CustomXMLNode node</param>
        /// <param name="errorName">string errorName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Add(NetOffice.OfficeApi.CustomXMLNode node, string errorName);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860763.aspx </remarks>
        /// <param name="node">NetOffice.OfficeApi.CustomXMLNode node</param>
        /// <param name="errorName">string errorName</param>
        /// <param name="errorText">optional string ErrorText = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Add(NetOffice.OfficeApi.CustomXMLNode node, string errorName, object errorText);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CustomXMLValidationError>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.CustomXMLValidationError> GetEnumerator();

        #endregion
    }
}
