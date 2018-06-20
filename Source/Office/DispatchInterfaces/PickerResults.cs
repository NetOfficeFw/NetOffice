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
    /// DispatchInterface PickerResults 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864136.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C03E5-0000-0000-C000-000000000046")]
    public interface PickerResults : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PickerResult>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.PickerResult this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865190.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        /// <param name="sIPId">optional string SIPId = </param>
        /// <param name="itemData">optional object itemData</param>
        /// <param name="subItems">optional object subItems</param>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId, object itemData, object subItems);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        /// <param name="sIPId">optional string SIPId = </param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864663.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="displayName">string displayName</param>
        /// <param name="type">string type</param>
        /// <param name="sIPId">optional string SIPId = </param>
        /// <param name="itemData">optional object itemData</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId, object itemData);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PickerResult>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.PickerResult> GetEnumerator();

        #endregion
    }
}
