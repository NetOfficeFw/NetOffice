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
    /// DispatchInterface PickerFields 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862500.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C03E1-0000-0000-C000-000000000046")]
    public interface PickerFields : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PickerField>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.PickerField this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862793.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PickerField>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.PickerField> GetEnumerator();

        #endregion
    }
}
