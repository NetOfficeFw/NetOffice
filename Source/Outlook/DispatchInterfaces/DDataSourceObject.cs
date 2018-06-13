using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface DDataSourceObject 
	/// SupportByVersion Outlook, 10
	/// </summary>
	[SupportByVersion("Outlook", 10)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("11858B51-DE06-494E-915A-6CCEF17F7CB6")]
    [CoClassSource(typeof(NetOffice.OutlookApi.DataSourceObject))]
    public interface DDataSourceObject : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 10), ProxyResult]
		object OutlookItem { get; set; }

		#endregion

	}
}
