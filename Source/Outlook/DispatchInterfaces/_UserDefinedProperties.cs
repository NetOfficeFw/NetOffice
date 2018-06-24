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
	/// DispatchInterface _UserDefinedProperties 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Outlook", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("00063047-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.UserDefinedProperties))]
    public interface _UserDefinedProperties : ICOMObject, IEnumerableProvider<NetOffice.OutlookApi._UserDefinedProperty>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868291.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860973.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867121.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869369.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866027.aspx </remarks>
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
		NetOffice.OutlookApi._UserDefinedProperty this[object index] { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869539.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlUserPropertyType type</param>
		/// <param name="displayFormat">optional object displayFormat</param>
		/// <param name="formula">optional object formula</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.UserDefinedProperty Add(string name, NetOffice.OutlookApi.Enums.OlUserPropertyType type, object displayFormat, object formula);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869539.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlUserPropertyType type</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.UserDefinedProperty Add(string name, NetOffice.OutlookApi.Enums.OlUserPropertyType type);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869539.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="type">NetOffice.OutlookApi.Enums.OlUserPropertyType type</param>
		/// <param name="displayFormat">optional object displayFormat</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.UserDefinedProperty Add(string name, NetOffice.OutlookApi.Enums.OlUserPropertyType type, object displayFormat);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862142.aspx </remarks>
		/// <param name="name">string name</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.UserDefinedProperty Find(string name);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866074.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Remove(Int32 index);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869393.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Refresh();

        #endregion

        #region IEnumerable<NetOffice.OutlookApi._UserDefinedProperty>

        /// <summary>
        /// SupportByVersion Outlook, 12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OutlookApi._UserDefinedProperty> GetEnumerator();

        #endregion
    }
}
