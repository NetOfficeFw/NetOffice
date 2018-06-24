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
	/// DispatchInterface _OrderFields 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Outlook", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("0006309A-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.OrderFields))]
    public interface _OrderFields : ICOMObject, IEnumerableProvider<NetOffice.OutlookApi._OrderField>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863629.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866929.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869189.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869720.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863107.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OutlookApi._OrderField this[object index] { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868461.aspx </remarks>
		/// <param name="propertyName">string propertyName</param>
		/// <param name="isDescending">optional object isDescending</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.OrderField Add(string propertyName, object isDescending);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868461.aspx </remarks>
		/// <param name="propertyName">string propertyName</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.OrderField Add(string propertyName);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869958.aspx </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Remove(object index);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861631.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void RemoveAll();

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868734.aspx </remarks>
		/// <param name="propertyName">string propertyName</param>
		/// <param name="index">object index</param>
		/// <param name="isDescending">optional object isDescending</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.OrderField Insert(string propertyName, object index, object isDescending);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868734.aspx </remarks>
		/// <param name="propertyName">string propertyName</param>
		/// <param name="index">object index</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.OrderField Insert(string propertyName, object index);

        #endregion

        #region IEnumerable<NetOffice.OutlookApi._OrderField>

        /// <summary>
        /// SupportByVersion Outlook, 12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OutlookApi._OrderField> GetEnumerator();

        #endregion
    }
}
