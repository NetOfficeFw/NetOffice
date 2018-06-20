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
    /// DispatchInterface PictureEffects 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864059.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Office", 14, 15, 16), HasIndexProperty(IndexInvoke.Property, "Item")]
	[TypeId("000C03D2-0000-0000-C000-000000000046")]
    public interface PictureEffects : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        NetOffice.OfficeApi.PictureEffect this[Int32 index] { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861170.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        Int32 Count { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861830.aspx </remarks>
        /// <param name="effectType">NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType</param>
        /// <param name="position">optional Int32 Position = -1</param>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PictureEffect Insert(NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType, object position);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861830.aspx </remarks>
        /// <param name="effectType">NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.PictureEffect Insert(NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862552.aspx </remarks>
        /// <param name="index">optional Int32 Index = -1</param>
        [SupportByVersion("Office", 14, 15, 16)]
        void Delete(object index);

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862552.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        void Delete();

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PictureEffect>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        new IEnumerator<NetOffice.OfficeApi.PictureEffect> GetEnumerator();

        #endregion
    }
}
