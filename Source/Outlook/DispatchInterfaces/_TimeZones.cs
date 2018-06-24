using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _TimeZones 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Outlook", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("000630FC-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.TimeZones))]
    public interface _TimeZones : ICOMObject, IEnumerableProvider<NetOffice.OutlookApi._TimeZone>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869536.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865050.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869616.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866936.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868840.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869781.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._TimeZone CurrentTimeZone { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OutlookApi._TimeZone this[object index] { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866193.aspx </remarks>
		/// <param name="sourceDateTime">DateTime sourceDateTime</param>
		/// <param name="sourceTimeZone">NetOffice.OutlookApi._TimeZone sourceTimeZone</param>
		/// <param name="destinationTimeZone">NetOffice.OutlookApi._TimeZone destinationTimeZone</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		DateTime ConvertTime(DateTime sourceDateTime, NetOffice.OutlookApi._TimeZone sourceTimeZone, NetOffice.OutlookApi._TimeZone destinationTimeZone);

        #endregion

        #region IEnumerable<NetOffice.OutlookApi._TimeZone>

        /// <summary>
        /// SupportByVersion Outlook, 12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OutlookApi._TimeZone> GetEnumerator();

        #endregion
    }
}
