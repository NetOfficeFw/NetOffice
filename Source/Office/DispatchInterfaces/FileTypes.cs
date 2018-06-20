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
    /// DispatchInterface FileTypes 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862049.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Value, EnumeratorInvoke.Property, "Office", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C036C-0000-0000-C000-000000000046")]
    public interface FileTypes : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.Enums.MsoFileType>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.Enums.MsoFileType this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864990.aspx </remarks>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861726.aspx </remarks>
        /// <param name="fileType">NetOffice.OfficeApi.Enums.MsoFileType fileType</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Add(NetOffice.OfficeApi.Enums.MsoFileType fileType);

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860597.aspx </remarks>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        void Remove(Int32 index);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.Enums.MsoFileType>

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.Enums.MsoFileType> GetEnumerator();

        #endregion
    }
}
