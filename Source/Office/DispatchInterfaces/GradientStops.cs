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
    /// DispatchInterface GradientStops 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861159.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C03C0-0000-0000-C000-000000000046")]
    public interface GradientStops : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.GradientStop>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.GradientStop this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864855.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861233.aspx </remarks>
        /// <param name="index">optional Int32 Index = -1</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Delete(object index);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861233.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Delete();

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        /// <param name="index">optional Int32 Index = -1</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Insert(Int32 rGB, Single position, object transparency, object index);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Insert(Int32 rGB, Single position);

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single transparency = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        void Insert(Int32 rGB, Single position, object transparency);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        /// <param name="index">optional Int32 Index = -1</param>
        /// <param name="brightness">optional Single brightness = 0</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void Insert2(Int32 rGB, Single position, object transparency, object index, object brightness);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        void Insert2(Int32 rGB, Single position);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        void Insert2(Int32 rGB, Single position, object transparency);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        /// <param name="index">optional Int32 Index = -1</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        void Insert2(Int32 rGB, Single position, object transparency, object index);

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.GradientStop>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.GradientStop> GetEnumerator();

        #endregion
    }
}
