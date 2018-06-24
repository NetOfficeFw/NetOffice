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
	/// DispatchInterface ItemProperties 
	/// SupportByVersion Outlook, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863357.aspx </remarks>
	[SupportByVersion("Outlook", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method, "Outlook", 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("000630A8-0000-0000-C000-000000000046")]
	public interface ItemProperties : ICOMObject, IEnumerableProvider<NetOffice.OutlookApi.ItemProperty>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868689.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863401.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865620.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862665.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864428.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.OutlookApi.ItemProperty this[object index] { get; }

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863099.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlUserPropertyType type</param>
		/// <param name="addToFolderFields">optional object addToFolderFields</param>
		/// <param name="displayFormat">optional object displayFormat</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.ItemProperty Add(string name, NetOffice.OutlookApi.Enums.OlUserPropertyType type, object addToFolderFields, object displayFormat);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863099.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlUserPropertyType type</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.ItemProperty Add(string name, NetOffice.OutlookApi.Enums.OlUserPropertyType type);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863099.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlUserPropertyType type</param>
		/// <param name="addToFolderFields">optional object addToFolderFields</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		NetOffice.OutlookApi.ItemProperty Add(string name, NetOffice.OutlookApi.Enums.OlUserPropertyType type, object addToFolderFields);

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865027.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		void Remove(Int32 index);

        #endregion

        #region IEnumerable<NetOffice.OutlookApi.ItemProperty>

        /// <summary>
        /// SupportByVersion Outlook, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Outlook", 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.OutlookApi.ItemProperty> GetEnumerator();

        #endregion
    }
}
