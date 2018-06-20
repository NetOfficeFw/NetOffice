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
    /// DispatchInterface PickerProperties 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861162.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C03E3-0000-0000-C000-000000000046")]
    public interface PickerProperties : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PickerProperty>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.PickerProperty this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861123.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863501.aspx </remarks>
        /// <param name="id">string id</param>
        /// <param name="value">string value</param>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoPickerField type</param>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PickerProperty Add(string id, string value, NetOffice.OfficeApi.Enums.MsoPickerField type);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863845.aspx </remarks>
        /// <param name="id">string id</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void Remove(string id);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PickerProperty>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.PickerProperty> GetEnumerator();

        #endregion
    }
}
