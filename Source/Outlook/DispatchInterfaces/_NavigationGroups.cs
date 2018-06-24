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
	/// DispatchInterface _NavigationGroups 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom, "Outlook", 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("000630EF-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.OutlookApi.NavigationGroups))]
    public interface _NavigationGroups : ICOMObject, IEnumerableProvider<NetOffice.OutlookApi._NavigationGroup>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867170.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._Application Application { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865787.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.Enums.OlObjectClass Class { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868713.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		NetOffice.OutlookApi._NameSpace Session { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869155.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866240.aspx </remarks>
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
		NetOffice.OutlookApi._NavigationGroup this[object index] { get; }

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865651.aspx </remarks>
		/// <param name="groupDisplayName">string groupDisplayName</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.NavigationGroup Create(string groupDisplayName);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868687.aspx </remarks>
		/// <param name="group">NetOffice.OutlookApi.NavigationGroup group</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		void Delete(NetOffice.OutlookApi.NavigationGroup group);

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868509.aspx </remarks>
		/// <param name="defaultFolderGroup">NetOffice.OutlookApi.Enums.OlGroupType defaultFolderGroup</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		NetOffice.OutlookApi.NavigationGroup GetDefaultNavigationGroup(NetOffice.OutlookApi.Enums.OlGroupType defaultFolderGroup);

        #endregion

        #region IEnumerable<NetOffice.OutlookApi._NavigationGroup>

        /// <summary>
        /// SupportByVersion Outlook, 12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Outlook", 12, 14, 15, 16)]
        [CustomEnumerator]
        new IEnumerator<NetOffice.OutlookApi._NavigationGroup> GetEnumerator();

        #endregion
    }
}
